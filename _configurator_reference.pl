#!/nslcm/system/bin/perl
# R(ICO) site
# Visio - MDF tab
#  !d4_oct! - fourth octect (4th octet is not currently in DB, needs added) /27
#  !v4_oct! /27
#  !e4_oct! /28
#  !s4_oct! (summary)StackList
#  !L4_oct! (loopback) /32

$|++;
sub devStrLen;
our $DEBUG = 0;

# TODO - utility to copy prod project, stack and form data from prod DB to dev, and generate a URL to run
# the script using that site info
#
# Sample URL:
# http://nslanwan-dev.uhc.com/configurator_dev.php?action=run&proj_id=WI105-Config&mailroute=WI105&id=477&projectid=207
# WI105 - mail route
# id=477 - is the id field in nslanwan.configurator_forms
# projectid=207 - column in nslanwan.projects
#
# DB copies are based on the mail route and encompass these tables:
# nslcm.siteList
# nslanwan.configurator_stacks
# nslanwan.configurator_forms
# nslanwan.projects
#
# All of the above entries for a mail route need to be copied from the prod DB to dev before configurator is run
#
# #############
#
# Add css, js and 'prettify' the output
#
# #############
#
### VERSION
our $VERSION = '16.0.3';
### MODULES
use CGI;
use CGI::Carp qw(fatalsToBrowser);
use Encode;
use DBI;
use File::Basename;
use Filesys::SmbClient;
use Filesys::SmbClientParser;
use POSIX;
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseExcel::SaveParser;
use strict;
use warnings;
no warnings ('uninitialized', 'substr');
use Data::Dumper;
use Encode qw(encode_utf8);
use HTTP::Request ();
use MIME::Base64;
use LWP::UserAgent;
use URI::Escape;
use File::Path qw(make_path);
use File::Find;
use File::Copy;
use Cwd;
use Env;
use Env qw(PATH HOME TERM);
use Env qw($SHELL @LD_LIBRARY_PATH);

#Environment variables need to be updated in dev in order to use the correct samba libraries
BEGIN {
	if ( $0 =~ m/dev/ ) {
		my $need = '/nslcm/system/qip/usr/lib:/nslcm/system/qip/jre/lib/sparc:/nslcm/system/sybase/OCS-15_0/lib:/nslcm/system/lib:/usr/lib/samba:/usr/lib/samba';
		my $newpath = '/nslcm/system/qip/jre/bin:/nslcm/system/qip/usr/bin:/nslcm/system/sybase/OCS-15_0/bin:/nslcm/system/bin:/nslcm/system/lib:/nslcm/sync:/usr/sbin:/usr/bin::/usr/bin:/usr/xpg4/bin:/usr/ccs/bin:/usr/ucb/bin:/usr/local/bin:/usr/sfw/bin:/usr/lib/samba:/usr/lib/samba';
		my $ld = $ENV{LD_LIBRARY_PATH};
		$ENV{'PATH'} = $newpath;
		$ENV{'USER'} = "nslcm";
		$ENV{'HOME'} = "/home/nslcm";
		if ( !$ld ) {
			$ENV{LD_LIBRARY_PATH} = $need;
		} elsif ( $ld !~ m#(^|:)\Q$need\E(:|$)# ) {
			$ENV{LD_LIBRARY_PATH} .= ':' . $need;
		} else {
			$need = "";
		}
		if ($need) {
			exec 'env', $^X, $0, @ARGV;
		}
	}
}

# Include files
our $DB_HOST;
our $DB_PASS;
our $SMB_USER;
our $SMB_PASS;
our $api_url;
require "/nslcm-ss/shared/settings_pl.inc";
require "/nslcm-ss/shared/brookes-routines.inc";

### SUB declarations (just so they can be called without &)

### Global SUB declarations
sub setAnchorValuesGlobal;
sub inputErrorCheckGlobal;
sub smbGet;
sub smbPut;
sub smbReadFile;
sub writeTemplate;
sub getstacks;
sub resolveAnchorValues;
sub VisioReIDTab;          # was VisioReidTab
sub VisioReadTabs;
sub VisioRenameTab;
sub VisioControlLayer;
sub writeRemoteSiteBuildChecklist;
sub deviceType;
sub writeRow;
sub prtout;                # replaces 'print' so debug messages can be output
sub xit;                   # single point of exit for the script - replaces any other 'exit' calls

### Legacy site models SUB declarations
sub setAnchorValuesLegacy;
sub getVisioTemplates;     # was updateVisioTemplates
sub inputErrorCheckLegacy;
sub wirelessAP;            # was handleWirelessAP
sub wirelessController;    # was handleWirelessController
sub mls_vdd;
sub mls_wdd;
sub mls_idd;
sub cis_wan_link;
sub mls_upl;
sub mls_upl_glbp;
sub cis_stl;
sub cis_stl_glbp;
sub mls_vnd;
sub writeSite;             # was a bunch of 'handle(somethingsomething)' subs
sub writeVisioPQ;          # was handlePsiteVisio, handleTriPsiteVisio, handleQsiteVisio
sub writeVisioS;           # was handleSsiteVisio
sub writeVisioXLM;         # was handleXLMsiteVisio
sub doswitches;
sub writeIPSummary;
sub writeCISummary;
sub writeEquipmentValidation;
sub setTadPorts;
sub writeSDWANcsv;

### New site models SUB declarations
sub setAnchorValuesNew;
sub getVisioTemplatesNew;
sub inputErrorCheckNew;
sub wirelessAPNew;
sub wirelessControllerNew;
sub mls_vddNew;
sub mls_wddNew;
sub mls_iddNew;
sub cis_man_link;
sub switch_upl;
sub mls_vndNew;
sub writeSiteNew;
sub writeVisioU;
sub writeVisioCF;          # was handlePsiteVisio, handleTriPsiteVisio, handleQsiteVisio
sub writeVisioDN;
sub writeVisioTGKH;
sub doswitchesNew;
sub writeIPSummaryNew;
sub writeCISummaryNew;
sub writeEquipmentValidationNew;
sub setTadPortsNew;
sub writeSDWANcsvNew;

prtout("DEBUG level: $DEBUG") if ($DEBUG);

### MAIN
my ( $id, $site_code, $environ, $msid );
( $id, $site_code, $environ, $msid  ) = split( /,/, $ARGV[0] );

# Environment notification
if ( $environ eq 'dev' or $environ eq '' ) {
	prtout( "=================================", "Entering Development Configurator",
			"Running Version: $VERSION",         "=================================" );
} else {
	prtout( "=================================", "Entering Production Configurator",
			"Running Version: $VERSION",         "=================================" );
}

# Database/SMB variables
our $DB_PORT   = '3306';
our $DB_USER   = 'configurator';
our $DB        = 'nslanwan';
our $DB_NSLCM  = 'nslcmdb';
our $SMB_DOM   = 'MS';
our $SMBCLIENT = '/usr/sfw/bin/smbclient';
my $site_model;

prtout("Host==>$DB_HOST");

# Set up DB/SMB connections
our $dbh = DBI->connect( "DBI:mysql:$DB:$DB_HOST:$DB_PORT", $DB_USER, $DB_PASS, { RaiseError => 1, AutoCommit => 1 } );
our $smb = new Filesys::SmbClient(
								   username  => $SMB_USER,
								   password  => $SMB_PASS,
								   workgroup => $SMB_DOM,
								   debug     => 0
);


my $usps;
# Mailcode to host mapping !#UPDATE!!!
my $query = $dbh->prepare( "SELECT USPS, DNS, DHCP, SNMP, NETFLOW, RADIUS, LOGGING FROM nslanwan.hostmappings" )
  or die "Query of mailcode-to-host mapping failed: " . $dbh->errstr;
$query->execute;
our ( %hostdns, %hostdhcp, %hostsnmp, %hostipt, %hostnetflow, %hostradius, %hostlogging, %attasn );
while ( my @HOSTMAP = $query->fetchrow_array ) {
	my @usps_data = split(/,/, $HOSTMAP[0]);
	foreach my $stmt (@usps_data){
	$usps = $stmt;
	$hostdns{$usps}     = $HOSTMAP[1];
	$hostdhcp{$usps}    = $HOSTMAP[2];
	$hostsnmp{$usps}    = $HOSTMAP[3];
	$hostnetflow{$usps} = $HOSTMAP[4];
	$hostradius{$usps}  = $HOSTMAP[5];
	$hostlogging{$usps} = $HOSTMAP[6];
	}
}

# File pathing - populated in setAnchorValues
our (
	  $SMB_TEMPLATE_DIR,    # templates
	  $SMB_FIN_DIR,         # base output dir
	  $OutputDir,           # output dir for user files - site_code-month-day
	  $SMB_ROOTDIR,         # base path for templates and finished files
	  $ROOTDIR,             # local directory
	  $SMB_SHARE_PATH,      # workaround for smbclient
	  $SVR_ROOTDIR,
);
$SMB_SHARE_PATH = "unpiox56pn.uhc.com/netsvcs";

# The old script set the stacks in various locations and depopulated the array in others
# Instead of doing that, populate a global, then copy/depopulate a temp array (if that's really necessary)
our @StackList = getstacks( $id, $site_code );
prtout( "Number of stacks: " . scalar(@StackList) );
prtout( "MSID: " . $msid );

# Set 'anchor' values
our %anchor;
my $sitemodel;

my $tstart = getTime();

setAnchorValuesGlobal( $id, $site_code, $environ );
my $sitetype = $anchor{'site_type'};

####################################################################################
#**********************SUBS FOR LEGACY SITE MODELS START HERE**********************#
####################################################################################

if ( $sitetype =~ /^(RDC|X|L|M|S|P|Q)$/ ) {
	setAnchorValuesLegacy( $site_code );

	# Set other site specific anchor values
	# Site branch logic. X,L,M, are handled by a generic route, S,P need special routines
	if ( $sitetype eq 'RDC' ) {
		writeSiteRDC(3);
	} elsif (    $sitetype eq 'P'
			or $sitetype eq 'S'
			or $sitetype eq 'Q'
			or $sitetype eq 'R')
	{
		writeSite();
	} elsif ( $sitetype eq 'M' ) {
		writeSite(4);
	} elsif ( $sitetype eq 'L' ) {
		writeSite(8);
	} elsif ( $sitetype eq 'X' ) {
		writeSite(16);
	}

my $tfinished = getTime();
my @ctrinfo = ($site_code,$tstart,$tfinished, $anchor{'proj_type'},$msid);
my $dbcounter = $dbh->prepare("INSERT INTO nslanwan.counter (SITECODE, DATECOUNTER,FINISHTIME, PRJ_TYPE, USER) values (?,?,?,?,?) ");
$dbcounter->execute(@ctrinfo);
xit(0);

sub setAnchorValuesLegacy {
	my $site_code = shift;
	my $sitetype = $anchor{'site_type'};
	$anchor{'site_name'} = $sitetype; #this is for RDC anchor values inside SDWAN consolidated template under writeSDWANcsvNew

	# Do initial input check
	inputErrorCheckLegacy();

	my $wantype = $anchor{'pri_circuit_type'};
	my $subfull = $anchor{'subrate'} * 1000;

	# FW variables sets
	if ( $anchor{'fw'} eq 'Y' ) {
		$anchor{'fw1_tad'} = 'FW1';
		$anchor{'fw2_tad'} = 'FW2';
		if ( $anchor{'site_type'} eq "L" ) {
			$anchor{'mls1_fw_config'} = smbReadFile("Model-L/mls1_fw_L.txt");
			$anchor{'mls2_fw_config'} = smbReadFile("Model-L/mls2_fw_L.txt");
		} elsif ( $anchor{'site_type'} eq "X" ) {
			$anchor{'mls1_fw_config'} = smbReadFile("Model-X/mls1_fw_X.txt");
			$anchor{'mls2_fw_config'} = smbReadFile("Model-X/mls2_fw_X.txt");
		} elsif ( $anchor{'site_type'} eq "M" ) {
			$anchor{'mls1_fw_config'} = smbReadFile("Model-M/mls1_fw_M.txt");
		}
	} else {
		$anchor{'fw1_tad'}        = 'empty';
		$anchor{'fw2_tad'}        = 'empty';
		$anchor{'mls1_fw_int'}    = " description reserved\r\n shutdown";
		$anchor{'mls1_fw_vl'}     = '!';
		$anchor{'mls1_fw_config'} = '!';
		$anchor{'mls2_fw_int'}    = " description reserved\r\n shutdown";
		$anchor{'mls2_fw_vl'}     = '!';
		$anchor{'mls2_fw_config'} = '!';
	}

	# Core Device names
	my $mdf_flr =
	  $anchor{'mdf_bldg'} . sprintf( "%02d", $anchor{'mdf_flrnumber'} );
	$anchor{'cis1_name'} = 'cis' . $site_code . $mdf_flr . 'a01';    # site types: all
	if ( $sitetype =~ /^[XLMS]$/ ) {                                 # site types: all but P and Q
		$anchor{'cis2_name'} = 'cis' . $site_code . $mdf_flr . 'a02';
	}
	if ( $sitetype =~ /^[XLM]$/ ) {                                  # site types: all but S, P and Q
		$anchor{'mls1_name'} = 'mls' . $site_code . $mdf_flr . 'a01';
		$anchor{'mls2_name'} = 'mls' . $site_code . $mdf_flr . 'a02';
	} elsif ( $sitetype eq 'RDC' ) {
	}         # rdc do not have stacks
	else {    # S, P and Q sites may need this set for the vgc config file
		$anchor{'mls1_name'} = ( split( /,/, $StackList[0] ) )[0];
		$anchor{'mls2_name'} = ( split( /,/, $StackList[0] ) )[0];
	}

	# Site type specific subnets
	my $oa = $anchor{'loop_oct1'};
	my $ob = $anchor{'loop_oct2'};
	my $oc = $anchor{'loop_oct3'};
	$anchor{'loop_subnet'} = "$oa.$ob.$oc";
	$oc++;
	$anchor{'data_subnet_1'} = "$oa.$ob.$oc";
	$anchor{'svr_subnet_1'}  = "$oa.$ob.$oc";

	# Convert data subnet to hex for Aruba vlan201 secondary IP
	my ($h1,$h2,$h3) = (sprintf("%02X", $oa), sprintf("%02X", $ob), sprintf("%02X", $oc));
	$anchor{'oct_to_hex'} = ("$h1:$h2:$h3");
	#prtout("$anchor{'oct_to_hex'}");

	# Voice subnet
	if ( $sitetype eq 'S' ) {
		$anchor{'data_subnet_1'} = "$oa.$ob.$oc";
		#$oc += 3;
		$oc++;
	} else {
		$anchor{'svr_subnet_1'} = "$oa.$ob.$oc";
		$oc++;
	}
	$anchor{'voice_subnet_1'} = "$oa.$ob.$oc";
	$anchor{'mdf_flr'} = substr( $anchor{'cis1_name'}, 8, 3 );

	# TAD
	$anchor{'tad1_name'}   = 'tad' . $site_code . $mdf_flr . 'a01';    # site types: all
	$anchor{'tad1_subnet'} = $anchor{'loop_subnet'} . '.160/30';
	$anchor{'tad1_ip'}     = $anchor{'loop_subnet'} . '.161';

	# WAAS
	$anchor{'waas1_name'}   = 'wae' . $site_code . $mdf_flr . 'a01';
	$anchor{'waas1_subnet'} = $anchor{'svr_subnet_1'} . '.152/30';
	$anchor{'waas1_ip'}     = $anchor{'svr_subnet_1'} . '.154';

	$anchor{'waas2_name'}   = 'wae' . $site_code . $mdf_flr . 'a02';
	$anchor{'waas2_subnet'} = $anchor{'svr_subnet_1'} . '.156/30';
	$anchor{'waas2_ip'}     = $anchor{'svr_subnet_1'} . '.158';

	# EAC Subnet
	if ( $sitetype eq 'S' or $sitetype eq 'P' or $sitetype eq 'Q' ) {
		$oc++;    # increment subnet for P & Q sites
		#$oc += 2 if ( $sitetype eq 'S' );    # S site updates by a total of 3
		$anchor{'eac_subnet_1'}        = "$oa.$ob.$oc";
		$anchor{'eac_subnet_wireless'} = '***DefnErr***';
	}

	# Router types
	# - If default router selected, determine router based on subrate
	# - If specific router is selected, use that setting
	if ( $anchor{'router_seltype'} =~ /^(4321|4331|4351|4451|4461|3945E|3945|2951|ASR|C8200-1N-4T|C8300-1N1S|C8300-2N2S|C8500-12X4QC)$/ ) {
		$anchor{'router_type'} = $anchor{'router_seltype'};
	} elsif ( $sitetype eq 'X' or $sitetype eq 'L' or $sitetype eq 'M' ) {    #XLM Default router code due to 4321 lack of ports.
		if ( $anchor{'pri_circuit_type'} =~ /^(T1|2xMLPPP|3xMLPPP|4xMLPPP|2xMLPPP-E1|3xMLPPP-E1|4xMLPPP-E1)$/ ) {
			$anchor{'router_type'} = '4351';
		} elsif ( $anchor{'pri_circuit_type'} =~ /^(Metro|MPLS)_Ethernet$/ ) {
			my $subrate = $anchor{'subrate'};
			if ( $subrate >= 50000 and $subrate < 200000 ) {                  #
				$anchor{'router_type'} = '4351';
			} elsif ( $subrate >= 200000 and $subrate <= 1000000 ) {
				$anchor{'router_type'} = '4451';
			} else {
				$anchor{'router_type'} = '4351';
			}
		} else {
			$anchor{'router_type'} = '4451';
		}
	} else {    # Set to default values
		if ( $anchor{'pri_circuit_type'} =~ /^(T1|2xMLPPP|3xMLPPP|4xMLPPP|2xMLPPP-E1|3xMLPPP-E1|4xMLPPP-E1)$/ ) {

			#			$anchor{'router_type'} = '2951';
			$anchor{'router_type'} = '4321';    # future standard - noted on 2017-04-20
		} elsif ( $anchor{'pri_circuit_type'} =~ /^(Metro|MPLS)_Ethernet$/ ) {
			my $subrate = $anchor{'subrate'};
			if ( $subrate >= 50000 and $subrate < 200000 ) {    #
				$anchor{'router_type'} = '4351';
			} elsif ( $subrate >= 200000 and $subrate <= 1000000 ) {
				$anchor{'router_type'} = '4451';
			} else {

				#				$anchor{'router_type'} = '2951';
				$anchor{'router_type'} = '4321';
			}
		} else {
			$anchor{'router_type'} = '4451';
		}
	}

	# P site uses 3945 (default) or 2951 (Tricare)
	# Adding 4451 for 'P' (pico) sites
	if ( $sitetype eq 'P' ) {
		$anchor{'router_type'} = 4331;    # new P site default 07-16-2020

		#		$anchor{'router_type'} = 3945;                                         # old default for P site
		$anchor{'router_type'} = 2951
		  if ( $anchor{'uhgdivision'} eq 'TRICARE' );    # tricare only
	}

	# Site wildcard mask
	my %wcmask = (
				   'P', '0.0.3.255',  'S', '0.0.7.255',  'Q',   '0.0.3.255', 'M', '0.0.15.255',
				   'L', '0.0.31.255', 'X', '0.0.63.255', 'RDC', '0.0.63.255'
	);
	if ( defined $wcmask{$sitetype} ) {
		$anchor{'wildcard_mask'} = $wcmask{$sitetype};
	} else {
		prtout( "Error: Selected site type '$sitetype' not currently supported in this version of Configurator.",
				"Please contact the developers concerning the availability of running Configurator with this model." );
		xit(1);
	}

		# Remove was staements from Q site model WAN
	if ( $sitetype eq 'Q' ) {
		if ( $wantype eq 'T1' or $wantype eq 'E1' ) {
			$anchor{'remove_wae'} = smbReadFile("Model-Q/T1E1_wae.txt");
		} elsif ( $wantype eq "Metro_Ethernet" || $wantype eq "MPLS_Ethernet" ) {
			$anchor{'remove_wae'} = smbReadFile("Model-Q/MetroE_wae.txt");
		} else {
			$anchor{'remove_wae'} = smbReadFile("Model-Q/MLPPP_wae.txt");
		}
	}
	# Visio template updates
	getVisioTemplates($sitetype);

	# Site branch logic. X,L,M are handled by a generic route, S,P need special routines
	if ( $sitetype eq 'M' ) {
		if ( $anchor{'ipt'} eq 'Y' ) {
			$anchor{'spare_amt'}   = 1;
			$anchor{'spare_name1'} = '9300_spare';
			$anchor{'spare_name2'} = 'empty';
		} else {
			$anchor{'spare_amt'}   = 0;
			$anchor{'spare_name1'} = 'empty';
			$anchor{'spare_name2'} = 'empty';
		}
	} elsif ( $sitetype eq 'L' ) {
		$anchor{'spare_amt'}   = 2;
		$anchor{'spare_name1'} = '9300_spare';
		$anchor{'spare_name2'} = '9300_spare';
	} elsif ( $sitetype eq 'X' ) {
		$anchor{'spare_amt'}   = 2;
		$anchor{'spare_name1'} = '9300_spare';
		$anchor{'spare_name2'} = '9300_spare';
	}
}

sub getVisioTemplates {    # was updateVisioTemplates
return unless ($anchor{'proj_type'} eq 'build');
	my $sitetype = shift;
	prtout("Downloading Visio template for site type '$sitetype'\n");
	if ( $sitetype eq 'RDC' ) {
		smbGet( "$SMB_TEMPLATE_DIR/Visio/rdc/master-rdc-template.vdx", "$ROOTDIR/master-rdc-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/Visio/rdc/master-rdc-template-clean.vdx", "$ROOTDIR/master-rdc-template-clean.vdx" );
	} elsif ( $sitetype eq 'X' or $sitetype eq 'L' or $sitetype eq 'M' ) {
		smbGet( "$SMB_TEMPLATE_DIR/Visio/generic-xlm/master-xlm-template.vdx",       "$ROOTDIR/master-xlm-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/Visio/generic-xlm/master-xlm-template-clean.vdx", "$ROOTDIR/master-xlm-template-clean.vdx" );
	} elsif ( $sitetype eq 'S' ) {
		smbGet( "$SMB_TEMPLATE_DIR/Visio/s_site/master-s-template.vdx",       "$ROOTDIR/master-s-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/Visio/s_site/master-s-template-clean.vdx", "$ROOTDIR/master-s-template-clean.vdx" );
	} elsif ( $sitetype eq 'P' ) {
		smbGet( "$SMB_TEMPLATE_DIR/Visio/p_site/master-p-template.vdx",       "$ROOTDIR/master-p-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/Visio/p_site/master-p-template-clean.vdx", "$ROOTDIR/master-p-template-clean.vdx" );
	} elsif ( $sitetype eq 'Q' ) {
		smbGet( "$SMB_TEMPLATE_DIR/Visio/q_site/master-q-template.vdx",       "$ROOTDIR/master-q-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/Visio/q_site/master-q-template-clean.vdx", "$ROOTDIR/master-q-template-clean.vdx" );
	} else {
		prtout( "Unknown site type! Please check the site type and resubmit, or contact the developer ",
				" if the site type is correct." );
		exit;
	}
}

sub inputErrorCheckLegacy {
	my $numstack = scalar(@StackList);
	my $sitetype = $anchor{'site_type'};
	my $firewall = $anchor{'fw'};
	my $xsubnet  = $anchor{'xsubnet'};
	if ( $sitetype eq 'S' and $firewall eq 'Y' ) {
		prtout(
				"Currently, Configurator does not support a Small site with a firewall.",
				"You will need to manually configure the firewall aspect of the Small site.",
				"Please discuss with Steve or Dan if you require a further explanation",
				"Site Type: $sitetype",
				"Firewall Check: $firewall"
		);
		xit(1);
	}
	if ( $sitetype eq 'P' and $firewall eq 'Y' ) {
		prtout(
"Currently, Configurator does not support a Pico site with a firewall. You will need to manually configure the firewall aspect of the Pico site.",
			"Please discuss with Steve or Dan if you require a further explanation",
			"Site Type: $sitetype",
			"Firewall Check: $firewall"
		);
		xit(1);
	}
	if ( $sitetype eq 'L' and $numstack > 8 ) {
		prtout(
			  "Currently, Configurator does not support a Large /19 site with more than 8 stacks due to IP addressing constraints.",
			  "Please discuss with Steve or Dan if you require a further explanation",
			  "Site Type: $sitetype",
			  "Firewall Check: $numstack"
		);
		xit(1);
	}
	if (     $sitetype eq 'M'
		 and $numstack > 3
		 and ( $firewall eq 'Y' or $xsubnet > 0 ) )
	{
		prtout(
"Currently, Configurator does not support a Medium site with 4 stacks and a FW or Extra Server subnets due to IP addressing constraints. You will probably need to explore and additional IP address block, run Configurator with 3 stacks, then manually adjust the outputs for the fourth stack.",
			"Please discuss with Steve or Dan if you require a further explanation",
			"Site Type: $sitetype",
			"Firewall Check: $numstack",
			"Firewall Check: $firewall",
			"Extra Server Subnets: $xsubnet"
		);
		xit(1);
	}
	if ( $sitetype eq 'S' and $firewall eq 'Y' ) {
		prtout(
"Currently, Configurator does not support a Small site with a FW. You will probably need to run a normal S site configuration and manually configure for a Firewall",
			"Site Type: $sitetype",
			"Firewall Check: $firewall"
		);
		xit(1);
	}
	if ( $anchor{'uhgdivision'} eq 'TRICARE' ) {
		if ( $sitetype =~ /^(?:X|L|M|S)$/ ) {
			prtout(
"Development for the TriCare configuration is in progress. Please check with Dan Bartlett or Steve Groebe for current status." );
		}
		xit(1);
	}
	if ( $sitetype eq 'S' and $anchor{'router_seltype'} eq '3945E' ) {
		prtout(
				"Currently, Configurator does not support a Small site with a 3945E Router selected.",
				"Please discuss with Steve or Dan if you require a further explanation.",
				"Site Type: $sitetype",
				"Router Check: $anchor{'router_seltype'}"
		);
		xit(1);
	}
}

sub wirelessAP {    # was handleWirelessAP
	my $sitetype = $anchor{'site_type'};
	$anchor{'flex_controllers'} = '!';
	if ( $anchor{'region'} eq 'USA' ) {
		if ( $sitetype =~ /^(?:X|L|M)$/ ) {
			$anchor{'wlc1_mgmt_ip'} = $anchor{'svr_subnet_1'} . '.9';    # local 5500 controller
			if( $anchor{'wireless_region'} eq 'EAST' ){ # Sets US east tertiary controller
				$anchor{'wlc_ter_ip'} = '10.141.58.144';
				$anchor{'wlc_ter_name'}    = 'wlcMN053bkpa21';
			}
			elsif( $anchor{'wireless_region'} eq 'WEST' ){ # Sets US west tertiary controller
				$anchor{'wlc_ter_ip'} = '10.141.62.149';
				$anchor{'wlc_ter_name'}    = 'wlcMN011bkpa21';
			}
		} elsif ( $sitetype =~ /^(?:S|P|Q)$/ ) {
				$anchor{'wlc1_mgmt_ip'} = '<FILL_IN_' . $anchor{'wireless_region'} . '_IP>';
				$anchor{'wlc1_name'} = '<FILL_IN_' . $anchor{'wireless_region'} . '_NAME>';
				if( $anchor{'wireless_region'} eq 'WEST' ) { # Sets west flex controllers
					$anchor{'wlc_ter_ip'} = '10.141.58.144';
					$anchor{'wlc_ter_name'}    = 'wlcMN053bkpa21';
					$anchor{'flex_controllers'} = smbReadFile("Wireless/AP_flex_west.txt");
				}
				else { # Sets east flex controllers
					$anchor{'wlc_ter_ip'} = '10.141.62.149';
					$anchor{'wlc_ter_name'}    = 'wlcMN011bkpa21';
					$anchor{'flex_controllers'} = smbReadFile("Wireless/AP_flex_east.txt");
				}
		}
	} else {
		if ( $sitetype =~ /^(?:X|L|M)$/ ) {
			$anchor{'wlc1_mgmt_ip'} = $anchor{'svr_subnet_1'} . '.9';    # local 5500 controller
			$anchor{'wlc_ter_ip'} = '10.177.22.9'; # Sets tertiary for all non-US
			$anchor{'wlc_ter_name'}    = 'wlcMN053bkpa01';
		} elsif ( $sitetype =~ /^(?:S|P|Q)$/ ) {
			$anchor{'wlc1_mgmt_ip'} = '<FILL_IN_' . $anchor{'wireless_region'} . '_IP>';
			$anchor{'wlc1_name'} = '<FILL_IN_' . $anchor{'wireless_region'} . '_NAME>';
			$anchor{'wlc_ter_ip'} = '10.141.62.144'; # Sets tertiary controller the same for all flex sites
			$anchor{'wlc_ter_name'}    = 'wlcMN011bkpa02';
			$anchor{'flex_controllers'} = smbReadFile("Wireless/AP_flex_east.txt"); # Sets international flex controllers
		}
	}
	if ( $anchor{'wlan'} eq 'Y' ) {
		prtout("Updating Wireless AP list file");
		writeTemplate( "Wireless/AP_data.txt", $anchor{'site_code'} . '-WAPs.txt' );
	}
	# AP placement Attestation
	writeTemplate( "Wireless/Attestation/AP Placement Attestation.xml", $anchor{'site_code'} . ' - AP Placement Attestation.doc' );
}

sub wirelessController{                                                 # was handleWirelessController
	return unless ( $anchor{'wlan'} eq 'Y' );
	my $sitetype = $anchor{'site_type'};
	if ( $sitetype eq 'L' or $sitetype eq 'X' ) {
		$anchor{'wireless_lag'} = 'lag disable';
	} elsif ( $sitetype eq 'M' ) {
		$anchor{'wireless_lag'} = 'lag enable';
	}
	prtout("Updating Wireless Controller Configuration");
	my $sitemask = '255.255.248.0';                                      # default
	if ( $sitetype eq 'X' ) {
		$sitemask                = '255.255.192.0';
		$anchor{'wlc1_mls_int'}  = 'GigabitEthernet5/0/9';
		$anchor{'wlc1_mls_int1'} = 'TenGigabitEthernet1/0/23';
		$anchor{'wlc1_mls_int2'} = 'TenGigabitEthernet1/0/24';
		$anchor{'wlc2_mls_int1'} = 'TenGigabitEthernet1/0/23';
		$anchor{'wlc2_mls_int2'} = 'TenGigabitEthernet1/0/24';
	} elsif ( $sitetype eq 'L' ) {
		$sitemask                = '255.255.224.0';
		$anchor{'wlc1_mls_int'}  = 'GigabitEthernet5/0/9';
		$anchor{'wlc1_mls_int1'} = 'TenGigabitEthernet1/0/23';
		$anchor{'wlc1_mls_int2'} = 'TenGigabitEthernet1/0/24';
		$anchor{'wlc2_mls_int1'} = 'TenGigabitEthernet1/0/23';
		$anchor{'wlc2_mls_int2'} = 'TenGigabitEthernet1/0/24';
	} elsif ( $sitetype eq 'M' ) {
		$sitemask                = '255.255.240.0';
		$anchor{'wlc1_mls_int1'} = 'TenGigabitEthernet1/1/8';
		$anchor{'wlc1_mls_int2'} = 'TenGigabitEthernet2/1/8';
		$anchor{'wlc2_mls_int1'} = 'TenGigabitEthernet1/1/7';
		$anchor{'wlc2_mls_int2'} = 'TenGigabitEthernet2/1/7';
	}

	# determine FW subnet as last data /24 then calculate wireless subnet
	( my $mask_octet1, my $mask_octet2, my $mask_octet3, my $mask_octet4 ) =
	  split( /\./, $sitemask );
	( my $wl_octet1, my $wl_octet2, my $wl_octet3 ) =
	  split( /\./, $anchor{'loop_subnet'} );
	$anchor{'eac_subnet_wireless'} = '***DefnErr***';
	if ( $sitetype eq 'X' ) {    # X sites receive /22 subnets for wireless
		my $eac_octet3 = $wl_octet3 + ( ( 254 - $mask_octet3 ) / 2 );
		my $data_octet3 = $eac_octet3 - 7;
		$anchor{'wlan_subnet_i'}          = $wl_octet1 . '.' . $wl_octet2 . '.' . $data_octet3;
		$anchor{'wlan_subnet_e'}          = $wl_octet1 . '.' . $wl_octet2 . '.' . $eac_octet3;
		$anchor{'wlan_subnet_mask'}       = '255.255.252.0';
		$anchor{'wlan_subnet_mask_nexus'} = '/22';
	} elsif ( $sitetype eq 'L' or $sitetype eq 'M' ) {    # L & M sites receive /24 subnets for wireless
		my $data_octet3 = $wl_octet3 + ( ( 254 - $mask_octet3 ) / 2 ) - 1;
		my $eac_octet3  = $wl_octet3 + ( ( 254 - $mask_octet3 ) / 2 );
		$anchor{'wlan_subnet_i'}          = $wl_octet1 . '.' . $wl_octet2 . '.' . $data_octet3;
		$anchor{'wlan_subnet_e'}          = $wl_octet1 . '.' . $wl_octet2 . '.' . $eac_octet3;
		$anchor{'wlan_subnet_mask'}       = '255.255.255.0';
		$anchor{'wlan_subnet_mask_nexus'} = '/24';
	} else {
		$anchor{'wlan_subnet_i'}          = $wl_octet1 . '.' . $wl_octet2 . '.' . '***Bad-WLAN-EAC-Subnet';
		$anchor{'wlan_subnet_e'}          = $wl_octet1 . '.' . $wl_octet2 . '.' . '***Bad-WLAN-EAC-Subnet';
		$anchor{'wlan_subnet_mask'}       = '255.255.255.0';
		$anchor{'wlan_subnet_mask_nexus'} = '/24';
	}
	my $mdf_flr =
	  $anchor{'mdf_bldg'} . sprintf( "%02d", $anchor{'mdf_flrnumber'} );
	$anchor{'wlc1_name'} = 'wlc' . $anchor{'site_code'} . $mdf_flr . 'a01';
	$anchor{'wlc2_name'} = 'wlc' . $anchor{'site_code'} . $mdf_flr . 'a02';

	# Region data
	my $state = '';
	if ( $anchor{'region'} eq 'USA' ) {
		prtout("Region: USA");
		$state = $anchor{'state'};
		$state =~ tr/a-z/A-Z/;
		if ( !( defined $hostnetflow{$state} ) or $hostnetflow{$state} eq '' ) {
			prtout(
					"There appears to be a problem in looking up host information for the state selected below.",
					"State: $state",
					"Please verify the state abbreviation is correct on the NSLANWAN website and rerun Configurator.",
					"If there continues to be issues, please contact Steve or Dan."
			);
			xit(1);
		}
	} else {
		prtout("Region: non-USA: $anchor{'region'}");
		$state = $anchor{'region'};
	}

	# Temporary solution to WLC snmp limitation, to be updated once all old snmp are removed
	my @east = ('AL','AR','CT','DE','FL','GA','IA','IL','IN','KY','LA','MA','MD','ME','MI','MO','MS','NC','NH','NJ','NY','OH',
	'PA','RI','SC', 'TN', 'VA', 'VT', 'WA', 'WI', 'WV', 'Asia (India and West)','Pacific (East of India)','Europe','Canada-CST');
	my @west = ('AK','AZ','CA','CO','HI','ID','KS','MN','MT','ND','NE','NM','NV','OK','OR','SD','TX', 'UT','WA','WY');

	if($state ~~ @east){
		$anchor{'snmp_host1'} = "10.208.154.191";
		$anchor{'snmp_host2'} = "10.177.72.106";
		$anchor{'snmp_host3'} = "10.122.72.169";
		$anchor{'snmp_host4'} = "10.87.57.127";
		$anchor{'snmp_host5'} = "10.86.186.96";
		$anchor{'snmp_host6'} = "10.86.142.225";
		$anchor{'wlc_snmp_name1'} = "cpieast.uhc.com";
		$anchor{'wlc_snmp_name2'} = "apslp0722";
		$anchor{'wlc_snmp_name3'} = "apsls0208";
		$anchor{'wlc_snmp_name4'} = "rp000057185";
		$anchor{'wlc_snmp_name5'} = "rn000057183";
		$anchor{'wlc_snmp_name6'} = "rp000073778";
		}

	if($state ~~ @west){
		$anchor{'snmp_host1'} = "10.208.155.88";
		$anchor{'snmp_host2'} = "10.177.72.148";
		$anchor{'snmp_host3'} = "10.122.72.188";
		$anchor{'snmp_host4'} = "10.87.57.127";
		$anchor{'snmp_host5'} = "10.86.186.96";
		$anchor{'snmp_host6'} = "10.29.74.248";
		$anchor{'wlc_snmp_name1'} = "cpiwest.uhc.com";
		$anchor{'wlc_snmp_name2'} = "apslp0724";
		$anchor{'wlc_snmp_name3'} = "apsls0210";
		$anchor{'wlc_snmp_name4'} = "rp000057185";
		$anchor{'wlc_snmp_name5'} = "rn000057183";
		$anchor{'wlc_snmp_name6'} = "vp000054652";
		}

	$anchor{'mobility_group'} = $anchor{'city'} . '-' . $anchor{'site_code'};
	$anchor{'wlc_sysloc'} =
	  $anchor{'city'} . '_' . $anchor{'state'} . '-' . $anchor{'site_code'};

	# City name may contain spaces - remove them
	$anchor{'mobility_group'} =~ s/\s//g;
	$anchor{'wlc_sysloc'} =~ s/\s//g;


	#9800 Config
	my $wlc_nmbr = $anchor{'wlc_nmbr'};
	if ($anchor{'wlc_model'} eq "9800"){
		if ($sitetype =~ /^(L|X)$/){
			$anchor{'wlc_uplink_9800'} = smbReadFile("Wireless/wlc_uplink_9800_L.txt");
			$anchor{'wlc_uplink_sec_9800'} = smbReadFile("Wireless/wlc_uplink_sec_9800_L.txt");
			writeTemplate( "Wireless/CTRL_9800_Legacy_Site_Model.txt",   $anchor{'wlc1_name'} . '-9800.txt' );
			writeTemplate( "Wireless/CTRL_9800_sec_Legacy_Site_Model.txt",   $anchor{'wlc2_name'} . '-9800.txt' );
		}else{
			if ($wlc_nmbr == 1){
			$anchor{'wlc_uplink_9800'} = smbReadFile("Wireless/wlc_uplink_9800_M.txt");
			writeTemplate( "Wireless/CTRL_9800_Legacy_Site_Model.txt",   $anchor{'wlc1_name'} . '-9800.txt' );
			}else{
				$anchor{'wlc_uplink_9800'} = smbReadFile("Wireless/wlc_uplink_9800_M.txt");
				$anchor{'wlc_uplink_sec_9800'} = smbReadFile("Wireless/wlc_uplink_sec_9800_M.txt");
				writeTemplate( "Wireless/CTRL_9800_Legacy_Site_Model.txt",   $anchor{'wlc1_name'} . '-9800.txt' );
				writeTemplate( "Wireless/CTRL_9800_sec_Legacy_Site_Model.txt",   $anchor{'wlc2_name'} . '-9800.txt' );
			}
		}
	}else{
		writeTemplate( "Wireless/CTRL_5500.txt",           $anchor{'wlc1_name'} . '.txt' );
		writeTemplate( "Wireless/CTRL_Initial_Config.txt", $anchor{'wlc1_name'} . '-initial.txt' );
	}
}
# VLAN definitions
sub mls_vdd {
	( my $mls, my $sitetype ) = @_;
	my $fw = $anchor{'fw'};
	$anchor{'extra_subnets'} = 0 if !defined( $anchor{'extra_subnets'} );
	my $e = $anchor{'extra_subnets'};

	# data offset is the amount of /24 networks between the loop and data subnets
	my %dataoffset = ( 'M', [ ( $e + 2 .. $e + 5 ) ], 'L', [ ( $e + 2 .. $e + 9 ) ], 'X', [ ( $e + 2 .. $e + 23 ) ], );

	# voice offset is the amount of /24 networks between the loop and data subnets
	my %voiceoffset =
	  ( 'M', [ ( 8 .. 11 ) ], 'L', [ ( 16 .. 23 ) ], 'X', [ ( 32 .. 53 ) ], );

	# eac offset is the amount of /24 networks between the loop and eac subnets
	my %eacoffset = (    # these values descend
					  'M', [ reverse( 12 .. 15 ) ],
					  'L', [ reverse( 24 .. 31 ) ],
					  'X', [ reverse( 42 .. 63 ) ],
	);
	my ( @voiceoffset, @dataoffset, @eacoffset );
	foreach my $offset ( @{ $voiceoffset{$sitetype} } ) {
		my $os = $offset + $anchor{'loop_oct3'};
		push @voiceoffset, '10.' . $anchor{'loop_oct2'} . '.' . $os;
	}
	foreach my $offset ( @{ $dataoffset{$sitetype} } ) {
		my $os = $offset + $anchor{'loop_oct3'};
		push @dataoffset, '10.' . $anchor{'loop_oct2'} . '.' . $os;
	}
	foreach my $offset ( @{ $eacoffset{$sitetype} } ) {
		my $os = $offset + $anchor{'loop_oct3'};
		push @eacoffset, '10.' . $anchor{'loop_oct2'} . '.' . $os;
	}
	my (
		 @datalastoct,  @eaclastoct,   @datavlanpri, @eacvlanpri, @datapreempt, @eacpreempt, @voicelastoct,
		 @voicevlanpri, @voicepreempt, @standby,     @datavlanid, @voicevlanid, @eacvlanid
	);

	# Defaults
	@standby      = ( 1 .. 22 );
	@datavlanid   = ( 201 .. 222 );
	@voicevlanid  = ( 301 .. 322 );
	@eacvlanid    = ( 401 .. 422 );
	@datalastoct  = ( 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2 );
	@eaclastoct   = ( 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2 );
	@datavlanpri  = ( 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110 );
	@eacvlanpri   = ( 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110 );
	@voicelastoct = ( 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3 );
	@voicevlanpri = ( 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90 );

	if ( $mls eq '1' ) {    # winds up being the reverse of the above arrays
		@datalastoct  = reverse(@datalastoct);
		@eaclastoct   = reverse(@eaclastoct);
		@datavlanpri  = reverse(@datavlanpri);
		@eacvlanpri   = reverse(@eacvlanpri);
		@voicelastoct = reverse(@voicelastoct);
		@voicevlanpri = reverse(@voicevlanpri);
	}

	# Preempts
	if ( $mls eq '1' ) {
		@datapreempt = (
						 "standby 201 preempt\r\n ", "", "standby 203 preempt\r\n ", "", "standby 205 preempt\r\n ", "",
						 "standby 207 preempt\r\n ", "", "standby 209 preempt\r\n ", "", "standby 211 preempt\r\n ", "",
						 "standby 213 preempt\r\n ", "", "standby 215 preempt\r\n ", "", "standby 217 preempt\r\n ", "",
						 "standby 219 preempt\r\n ", "", "standby 221 preempt\r\n ", ""
		);
		@eacpreempt = (
						"standby 61 preempt\r\n ", "", "standby 63 preempt\r\n ", "", "standby 65 preempt\r\n ", "",
						"standby 67 preempt\r\n ", "", "standby 69 preempt\r\n ", "", "standby 71 preempt\r\n ", "",
						"standby 73 preempt\r\n ", "", "standby 75 preempt\r\n ", "", "standby 77 preempt\r\n ", "",
						"standby 79 preempt\r\n ", "", "standby 81 preempt\r\n ", ""
		);
		@voicepreempt = (
						  "", "standby 2 preempt\r\n ",  "", "standby 4 preempt\r\n ",  "", "standby 6 preempt\r\n ",
						  "", "standby 8 preempt\r\n ",  "", "standby 10 preempt\r\n ", "", "standby 12 preempt\r\n ",
						  "", "standby 14 preempt\r\n ", "", "standby 16 preempt\r\n ", "", "standby 18 preempt\r\n ",
						  "", "standby 20 preempt\r\n ", "", "standby 22 preempt\r\n "
		);
	} else {
		@datapreempt = (
						 "", "standby 202 preempt\r\n ", "", "standby 204 preempt\r\n ", "", "standby 206 preempt\r\n ",
						 "", "standby 208 preempt\r\n ", "", "standby 210 preempt\r\n ", "", "standby 212 preempt\r\n ",
						 "", "standby 214 preempt\r\n ", "", "standby 216 preempt\r\n ", "", "standby 218 preempt\r\n ",
						 "", "standby 220 preempt\r\n ", "", "standby 222 preempt\r\n "
		);
		@eacpreempt = (
						"", "standby 62 preempt\r\n ", "", "standby 64 preempt\r\n ", "", "standby 66 preempt\r\n ",
						"", "standby 68 preempt\r\n ", "", "standby 70 preempt\r\n ", "", "standby 72 preempt\r\n ",
						"", "standby 74 preempt\r\n ", "", "standby 76 preempt\r\n ", "", "standby 78 preempt\r\n ",
						"", "standby 80 preempt\r\n ", "", "standby 82 preempt\r\n "
		);
		@voicepreempt = (
						  "standby 1 preempt\r\n ",  "", "standby 3 preempt\r\n ",  "", "standby 5 preempt\r\n ",  "",
						  "standby 7 preempt\r\n ",  "", "standby 9 preempt\r\n ",  "", "standby 11 preempt\r\n ", "",
						  "standby 13 preempt\r\n ", "", "standby 15 preempt\r\n ", "", "standby 17 preempt\r\n ", "",
						  "standby 19 preempt\r\n ", "", "standby 21 preempt\r\n ", ""
		);
	}

	# Add definitions
	my $vlData = '';
	my $vlDef  = '';
	my $vlTemplate;
	my $idfct;

	# Unset currentstack so dynamic values can be used in the template
	my $tmpAnchor = delete( $anchor{'currentstack'} );

	# Data VLAN interfaces
	if ( $sitetype eq 'M' ) {
		$vlTemplate = smbReadFile("Generic-XLM-MLS/mls_M_vdd_data.txt");
	} else {
		$vlTemplate = smbReadFile("Generic-XLM-MLS/mls_vdd_data.txt");
	}
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $vlTmp = $vlTemplate;
		my $stack = ( split( /,/, $StackList[$ct] ) )[0];
		$vlTmp =~ s/\!vlanid\!/$datavlanid[$ct]/g;
		$vlTmp =~ s/\!currentstack\!/$stack/g;
		$vlTmp =~ s/\!data_subnet\!/$dataoffset[$ct]/g;
		$vlTmp =~ s/\!lastoct\!/$datalastoct[$ct]/g;
		$vlTmp =~ s/\!vlanpri\!/$datavlanpri[$ct]/g;
		$vlTmp =~ s/\!prempt\!/$datapreempt[$ct]/g;
		$vlTmp =~ s/\!dhcp_host\!/$anchor{'dhcp_host'}/g;
		$vlTmp =~ s/\!dhcp_host_nexus\!/$anchor{'dhcp_host_nexus'}/g;
		$vlDef .= "$vlTmp\r\n";
	}

	# Voice VLAN interfaces
	if ( $sitetype eq 'M' ) {
		$vlTemplate = smbReadFile("Generic-XLM-MLS/mls_M_vdd_voice.txt");
	} else {
		$vlTemplate = smbReadFile("Generic-XLM-MLS/mls_vdd_voice.txt");
	}
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		( my $stack, my $switchct ) = split( /,/, $StackList[$ct] );
		my $vlTmp = $vlTemplate;
		$vlTmp =~ s/\!vlanid\!/$voicevlanid[$ct]/g;
		$vlTmp =~ s/\!currentstack\!/$stack/g;
		$vlTmp =~ s/\!voice_subnet\!/$voiceoffset[$ct]/g;
		$vlTmp =~ s/\!lastoct\!/$voicelastoct[$ct]/g;
		$vlTmp =~ s/\!vlanpri\!/$voicevlanpri[$ct]/g;
		$vlTmp =~ s/\!standby\!/$standby[$ct]/g;
		$vlTmp =~ s/\!prempt\!/$voicepreempt[$ct]/g;
		$vlTmp =~ s/\!dhcp_host\!/$anchor{'dhcp_host'}/g;
		$vlTmp =~ s/\!dhcp_host_nexus\!/$anchor{'dhcp_host_nexus'}/g;
		$vlDef .= "$vlTmp\r\n";

		# These are used in templates that are read in other parts of the script
		$anchor{'currentstack'}         = $stack;
		$anchor{'current_voice_vlan'}   = $voicevlanid[$ct];
		$anchor{'current_data_vlan'}    = $datavlanid[$ct];
		$anchor{'current_eac_vlan'}     = $eacvlanid[$ct];
		$anchor{'current_data_subnet'}  = $dataoffset[$ct];
		$anchor{'current_voice_subnet'} = $voiceoffset[$ct];
		$anchor{'current_eac_subnet'}   = $eacoffset[$ct];

		# These are for the Visio answer file
		$idfct = $ct + 1;
		my $k = 'idf' . $idfct;    # keys for the below
		$anchor{ $k . '_flr' } = substr( $stack, 8,  3 );
		$anchor{ $k . '_rm' }  = substr( $stack, 11, 3 );
		$anchor{ $k . '_ds_1' }           = $dataoffset[$ct];
		$anchor{ $k . '_vs_1' }           = $voiceoffset[$ct];
		$anchor{ $k . '_es_1' }           = $eacoffset[$ct];
		$anchor{ $k . '_Data_vlan_1' }    = $datavlanid[$ct];
		$anchor{ $k . '_voice_vlan_1' }   = $voicevlanid[$ct];
		$anchor{ $k . '_EAC_vlan_1' }     = $eacvlanid[$ct];
		$anchor{ $k . '_dlo_mls' . $mls } = $datalastoct[$ct];
		$anchor{ $k . '_vlo_mls' . $mls } = $voicelastoct[$ct];
		$anchor{ $k . '_elo_mls' . $mls } = $eaclastoct[$ct];
		if ( $sitetype eq 'L' or $sitetype eq 'X' and $mls eq '1' ) {
			if ($anchor{'stack_vendor'} eq 'aruba') {
				writeTemplate( "Generic-XLM-Stacks/aruba/stk_$switchct" . '.txt', $stack . '.txt' );
			}else{
				writeTemplate( "Generic-XLM-Stacks/stk_$switchct" . '.txt', $stack . '.txt' );
			}
		}
		elsif ( $mls eq '1' ) {
			if ($anchor{'stack_vendor'} eq 'aruba') {
				writeTemplate( "Generic-XLM-Stacks/aruba/stk_$switchct" . '_channel.txt', $stack . '.txt' );
			} else{
				writeTemplate( "Generic-XLM-Stacks/stk_$switchct" . '_channel.txt', $stack . '.txt' );
			}
		}
	}

	# Unset currentstack (again, sheesh) so dynamic values can be used in the template
	$tmpAnchor = delete( $anchor{'currentstack'} );

	# EAC VLAN interfaces
	if ( $sitetype eq 'M' ) {
		$vlTemplate = smbReadFile("Generic-XLM-MLS/mls_M_vdd_eac.txt");
	} else {
		$vlTemplate = smbReadFile("Generic-XLM-MLS/mls_vdd_eac.txt");
	}
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $stack      = ( split( /,/, $StackList[$ct] ) )[0];
		my $vlTmp      = $vlTemplate;
		my $eacstandby = $ct + 61;
		$vlTmp =~ s/\!vlanid\!/$eacvlanid[$ct]/g;
		$vlTmp =~ s/\!currentstack\!/$stack/g;
		$vlTmp =~ s/\!eac_subnet\!/$eacoffset[$ct]/g;
		$vlTmp =~ s/\!lastoct\!/$eaclastoct[$ct]/g;
		$vlTmp =~ s/\!vlanpri\!/$eacvlanpri[$ct]/g;
		$vlTmp =~ s/\!prempt\!/$eacpreempt[$ct]/g;
		$vlTmp =~ s/\!dhcp_host\!/$anchor{'dhcp_host'}/g;
		$vlTmp =~ s/\!dhcp_host_nexus\!/$anchor{'dhcp_host_nexus'}/g;
		$vlTmp =~ s/\!standby_eac\!/$eacstandby/g;
		$vlDef .= "$vlTmp\r\n";
	}

	# The answer file requires blank entries for vlans that aren't configured, so create them if necessary
	$idfct++;    # increment it so we're at the next unused value
	for ( ; $idfct <= 22 ; $idfct++ ) {
		$anchor{ 'idf' . $idfct . '_flr' }  = '';
		$anchor{ 'idf' . $idfct . '_ds_1' } = '';
		$anchor{ 'idf' . $idfct . '_vs_1' } = '';
	}

	# Replace the value that was deleted just prior to the above loop for the EAC VLANs.
	$anchor{'currentstack'} = $tmpAnchor;
	return $vlDef;
}

sub mls_wdd {
	( my $mlswireless, my $sitetype ) = @_;
	my $wddDef = '';
	if ( $sitetype eq 'M' ) {
		if ( $mlswireless == 1 and $anchor{'wlc_model'} eq '9800') {
			$wddDef = smbReadFile("Wireless/MLS1_wireless_M_9800.txt");
		}elsif( $mlswireless == 2 and $anchor{'wlc_model'} eq '9800'){
			$wddDef = smbReadFile("Wireless/MLS1_wireless_M_with_sec_9800.txt");
		}else {
			$wddDef = smbReadFile("Wireless/MLS1_wireless_M.txt");
		}
	}else{
			if ( $mlswireless == 1 ){
				if ($anchor{'wlc_model'} eq '9800'){
					$wddDef = smbReadFile("Wireless/MLS1_wireless_9800.txt");
				}else{
					$wddDef = smbReadFile("Wireless/MLS1_wireless.txt");
				}
			}elsif( $mlswireless == 2 ){
				if ($anchor{'wlc_model'} eq '9800'){
					$wddDef = smbReadFile("Wireless/MLS2_wireless_9800.txt");
				}else{
					$wddDef = smbReadFile("Wireless/MLS2_wireless.txt");
				}
		 }
	}
	return $wddDef;
}

sub mls_idd {
	my $sitetype = shift;
	my ( @intconf, @intconf1, @intconf2, @trunkvlans, @portchannel, @portchannel_M );
	if ( $sitetype eq 'M' ) {
		@intconf  = ( '1/1/1', '1/1/2', '1/1/3', '1/1/4' );
		@intconf1 = ( '1/1/1', '1/1/2', '1/1/3', '1/1/4' );
		@intconf2 = ( '2/1/1', '2/1/2', '2/1/3', '2/1/4' );
	} elsif ( $sitetype eq 'L' or $sitetype eq 'X' ) {
		@intconf = (
					 '1/0/1',  '1/0/2',  '1/0/3',  '1/0/4',  '1/0/5',  '1/0/6',  '1/0/7',  '1/0/8',
					 '1/0/9',  '1/0/10', '1/0/11', '1/0/12', '1/0/13', '1/0/14', '1/0/15', '1/0/16',
					 '1/0/17', '1/0/18', '1/0/19', '1/0/20', '1/0/21', '1/0/22'
		);
	}
	for ( my $ct = 201 ; $ct <= 222 ; $ct++ ) {
		my $c2 = $ct + 100;
		my $c3 = $ct + 200;
		push @trunkvlans,  $ct . ',' . $c2 . ',' . $c3;    # eg '201,301,401'
		push @portchannel, $ct;
	}
	@portchannel_M = ( 1, 2, 3, 4 );
	my $intlimit;
	$intlimit = 4  if ( $sitetype eq 'M' );
	$intlimit = 8  if ( $sitetype eq 'L' );
	$intlimit = 16 if ( $sitetype eq 'X' );

	# read generic interface def file and add an interface for each stack
	my $iddOrig = smbReadFile( "Model-" . $sitetype . '/mls_idd_' . $sitetype . '.txt' );
	my $iddDef  = '';
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $iddTmp = $iddOrig;                            #make a copy of the original since symbols will be replaced multiple times
		my $idfct  = $ct + 1;
		my $stack  = ( split( /,/, $StackList[$ct] ) )[0];
		$iddTmp =~ s/\!intconf\!/$intconf[$ct]/g;
		$iddTmp =~ s/\!intconf1\!/$intconf1[$ct]/g;
		$iddTmp =~ s/\!intconf2\!/$intconf2[$ct]/g;
		$iddTmp =~ s/\!currentstack\!/$stack/g;
		$iddTmp =~ s/\!trunkvlans\!/$trunkvlans[$ct]/g;
		$iddTmp =~ s/\!portchannel\!/$portchannel[$ct]/g;
		$iddTmp =~ s/\!port-channel-m\!/$portchannel_M[$ct]/g;
		$iddTmp .= "\r\n";
		$iddDef .= $iddTmp;

		# Save values for Visio template
		if ( $sitetype eq 'L' or $sitetype eq 'X' ) {
			$anchor{ 'idf' . $idfct . '_mls_up' }   = $intconf[$ct];
			$anchor{ 'idf' . $idfct . '-1_mls_up' } = $intconf[$ct];
			$anchor{ 'idf' . $idfct . '-2_mls_up' } = $intconf[$ct];
		} elsif ( $sitetype eq 'M' ) {
			$anchor{ 'idf' . $idfct . '_mls_up' }   = $intconf[$ct];
			$anchor{ 'idf' . $idfct . '-1_mls_up' } = $intconf1[$ct];
			$anchor{ 'idf' . $idfct . '-2_mls_up' } = $intconf2[$ct];
		}
	}

	# Fill in the rest of the available interfaces with a shutdown statement
	for ( my $ct = scalar(@StackList) ; $ct < $intlimit ; $ct++ ) {
		if ( $sitetype eq 'M' ) {
			$iddDef .= 'interface GigabitEthernet' . $intconf[$ct] . "\r\n";
			$iddDef .= " shutdown\r\n!\r\n";
		}
	}
	return $iddDef;
}

sub cis_wan_link {
	( my $sitetype, my $routerNum, my $wanlink ) = @_;

	# There are around 40 different templates to pick, depending on the site type, wan link, router number and
	# sometimes router type, and the template files loosely correspond to these values. In the interests
	# of brevity I'm building the filename dynamically, but it's not the best thing to do. Not sure what the
	# final code will look like here. -zzz
	$wanlink =~ tr/A-Z/a-z/;
	if    ( $wanlink =~ /(\d)xmlppp-e1/ ) { $wanlink = $1 . 'e1_mlppp'; }
	elsif ( $wanlink =~ /(\d)xmlppp/ )    { $wanlink = $1 . 't1_mlppp'; }
	if    ( $wanlink !~ /(?:mpls|metro_ethernet)/ ) {
		$wanlink .= '_wan';
	} elsif ( $sitetype eq 'XLM' and $anchor{router_type} eq 'ASR' ) {
		$wanlink .= '_ASR';
	}
	my $readFile = "Generic-WAN/" . 'cis' . $routerNum . '_' . $wanlink . '.txt';
	if ( $routerNum eq '1' or $routerNum eq '2' ) {
		prtout( 'Writing CIS' . $routerNum . ' WAN Configuration' );
	} else {
		prtout("Could not identify router number for WAN configuration");
	}
	my $wanDef = smbReadFile($readFile);
	return $wanDef;
}

# The MLS configuration on the uplinks interfaces to the routers changes depending if an IPS device is
# installed or not. Configure a set of arrays that contain a sequential list of parameters for use in a
# generic VLAN definition.
sub mls_upl {
	( my $mls, my $sitetype, my $ips ) = @_;
	my $uplDef = '';
	my $readFile;
	if ( $sitetype eq 'S' ) {
		if ( $mls eq '24-port' ) {
			$mls = 1;
		}
		if ( $ips eq 'Y' ) {
			if ($anchor{'stack_vendor'} eq 'aruba'){
				$readFile = "Model-$sitetype/aruba" . '/stk' . $mls . '_ips.txt';
			}else{
				$readFile = "Model-$sitetype" . '/stk' . $mls . '_ips.txt';
			}
		} else {
			if ($anchor{'stack_vendor'} eq 'aruba'){
				$readFile = "Model-$sitetype/aruba" . '/stk' . $mls . '_stl.txt';
			}else{
				$readFile = "Model-$sitetype" . '/stk' . $mls . '_stl.txt';
			}
		}
	} else {
		if ( $ips eq 'Y' ) {
			$readFile = "Model-$sitetype" . '/mls' . $mls . '_ips_' . $sitetype . '.txt';
		} else {
			$readFile = "Model-$sitetype" . '/mls' . $mls . '_upl_' . $sitetype . '_sdwan.txt';
		}
	}
	$uplDef = smbReadFile($readFile);
	return $uplDef;
}

# The MLS configuration on the uplinks interfaces to the routers changes depending if an IPS device is
# installed or not. Configure a set of arrays that contain a sequential list of parameters for use in a
# generic VLAN definition.
sub mls_upl_glbp {
	( my $mls, my $sitetype, my $ips ) = @_;
	my $uplDef = '';
	my $readFile;
	if ( $sitetype eq 'S' ) {
		if ( $mls eq '24-port' ) {
			$mls = 1;
		}
		if ( $mls eq '3' or $mls eq '4' or $mls eq '5' ) {    # I think this just determines which template is read
			$mls = 2;
		}
		if ( $ips eq 'Y' ) {
			if ($anchor{'stack_vendor'} eq 'aruba'){
				$readFile = "Model-$sitetype/aruba" . '/stk' . $mls . '_ips_glbp.txt';
			}else{
				$readFile = "Model-$sitetype" . '/stk' . $mls . '_ips_glbp.txt';
			}
		} else {
			if ($anchor{'stack_vendor'} eq 'aruba'){
				$readFile = "Model-$sitetype/aruba" . '/stk' . $mls . '_stl_glbp.txt';
			}else{
				$readFile = "Model-$sitetype" . '/stk' . $mls . '_stl_glbp.txt';
			}
		}

	} else {
		if ( $ips eq 'Y' ) {
			$readFile = "Model-$sitetype" . '/mls' . $mls . '_ips_' . $sitetype . '.txt';
		} else {
			$readFile = "Model-$sitetype" . '/mls' . $mls . '_upl_' . $sitetype . '.txt';
		}
	}
	$uplDef = smbReadFile($readFile);
	return $uplDef;
}

# Determine the configuration for the Gig interface. This depends on whether in IPS is installed.
sub cis_stl {
	( my $cis, my $sitetype, my $ips ) = @_;
	my $stlDef = '';
	my $readFile;
	if ( $ips eq 'Y' ) {
		$readFile = "Model-" . $sitetype . '/cis' . $cis . '_ips.txt';
	} else {
		$readFile = "Model-" . $sitetype . '/cis' . $cis . '_stl.txt';
	}
	$stlDef = smbReadFile($readFile);
	return $stlDef;
}

# Determine the configuration for the Gig interface. This depends on whether in IPS is installed.
sub cis_stl_glbp {
	( my $cis, my $sitetype, my $ips ) = @_;
	my $stlDef = '';
	my $readFile;
	if (    $anchor{'router_type'} eq '3945E'
		 or $anchor{'router_type'} eq '3945'
		 or $anchor{'router_type'} eq '2951' )
	{
		$readFile = "Model-" . $sitetype . '/cis' . $cis . '_stl_glbp_ISR.txt';
	} else {
		$readFile = "Model-" . $sitetype . '/cis' . $cis . '_stl_glbp.txt';
	}
	$stlDef = smbReadFile($readFile);
	return $stlDef;
}

# Generate data, voice and EAC vlan names for each stack
sub mls_vnd {
	my $vlanDef = '';

	# Data
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $vl    = 200 + $ct + 1;
		my $datac = $ct + 1;
		$vlanDef .= 'vlan ' . $vl . "\r\n name IDF" . $datac . "_Data\r\n!\r\n";
	}

	# Voice
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $vl    = 300 + $ct + 1;
		my $datac = $ct + 1;
		$vlanDef .= 'vlan ' . $vl . "\r\n name IDF" . $datac . "_Voice\r\n!\r\n";
	}

	# EAC
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $vl    = 400 + $ct + 1;
		my $datac = $ct + 1;
		$vlanDef .= 'vlan ' . $vl . "\r\n name IDF" . $datac . "_EAC\r\n!\r\n";
	}
	chop($vlanDef);    #zzz why?
	return $vlanDef;
}

# Routine to output site data
sub writeSite {
		my $stacklimit = shift || 1;    # default, some site types will pass higher values

		# Site specific info
		my $sitetype = $anchor{'site_type'};
		my $stack    = substr( $anchor{'cis1_name'}, -11, 11 );
		my $wantype  = $anchor{'pri_circuit_type'};

		$anchor{'acl_tad'} = smbReadFile("Modules/acl_tad.txt");
		$anchor{'aruba_acl_tad'} = smbReadFile("Modules/aruba_acl_tad.txt");

		# Some of the templates have !currentstack! symbols, but that anchor value is not set
		# for some sites, so it needs to be set here.
		# Note: X, L and M sites set 'currentstack' in the mls_vdd sub
		if ( $sitetype eq 'P' or $sitetype eq 'Q' ) {

			# There will be only one 'stack' in a P site, which is a module in the CIS router
			# This means the 'stack' must be on the same floor as the router
			$anchor{'currentstack'} =
			'stk' . substr( $anchor{'cis1_name'}, -11, 11 );
		} elsif ( $sitetype eq 'S' ) {

			# S sites have only one stack so use the existing stack value
			$anchor{'currentstack'} = ( split( /,/, $StackList[0] ) )[0];
		}
		if ( $sitetype eq 'P' ) {
			#VGC interface from switch
			if ( $anchor{'vgcount'} > 0 ) {
				$anchor{'vgc_interface_uplink'} = smbReadFile("Model-P/stk1_stl_vgc.txt");

				$anchor{'vgc1_name'} = 'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01';
				$anchor{'vgc1_address'} = $anchor{'data_subnet_1'} . '.16';
				my $of = $anchor{'vgc1_name'} . '.txt';
				writeTemplate( "Misc/vgc_204XM_MDF_mls1.txt", $of );
			}
			else{
				$anchor{'vgc_interface_uplink'} = '';
			}

			$anchor{'stk_interface_uplink'} = smbReadFile("Model-P/stk1_stl.txt");

			# Hard coding the current_data_vlan in for P, Q and S sites, per Deven - 2016-09-30
			$anchor{'current_data_vlan'} = '201';



			if ($anchor{'stack_vendor'} eq 'aruba'){
				$anchor{'stk_interface_uplink'} = smbReadFile("Model-P/aruba/stk1_stl.txt");
				writeTemplate( "Model-P/aruba/stk_48p.txt", 'stk' . $stack . '.txt' );
			}else{
				$anchor{'stk_interface_uplink'} = smbReadFile("Model-P/stk1_stl.txt");
				writeTemplate( "Model-P/stk_48p.txt", 'stk' . $stack . '.txt' );
			}


			wirelessAP();
			$anchor{'tad_hosts'} = smbReadFile("Modules/pq_tad_hosts.txt");
		} elsif ( $sitetype eq 'Q' ) {

			# Hard coding the current_data_vlan in for P, Q and S sites, per Deven - 2016-09-30
			$anchor{'current_data_vlan'} = '201';
			writeTemplate( "Model-Q/stk_24p.txt", 'stk' . $stack . '.txt' );
			$anchor{'cis1_interface_stlink'} =
			cis_stl_glbp( 1, 'Q', $anchor{'ips'} );
			$anchor{'cis_wan_config'} = cis_wan_link( 'Q', 1, $wantype );
			if ( $wantype eq 'Metro_Ethernet' ) {
				writeTemplate( "Model-Q/cis1_base_metroE.txt", $anchor{'cis1_name'} . '.txt' );
			} else {
				writeTemplate( "Model-Q/cis1_base_glbp.txt", $anchor{'cis1_name'} . '.txt' );
			}
			wirelessAP();
			$anchor{'tad_hosts'} = smbReadFile("Modules/pq_tad_hosts.txt");
		} elsif ( $sitetype eq 'S' ) {


				( $stack, my $switchtype ) = split( /,/, $StackList[0] );
				if (    $anchor{'router_type'} eq '3945E'
					or $anchor{'router_type'} eq '3945'
					or $anchor{'router_type'} eq '2951' )
				{
					if ( $anchor{'fw'} eq 'Y' ) {
						$anchor{'cis1_fw_int_gen'} = smbReadFile("Model-S/cis1_fw_glbp_ISR.txt");
						$anchor{'cis2_fw_int_gen'} = smbReadFile("Model-S/cis2_fw_glbp_ISR.txt");
						$anchor{'stk1_fw_int_gen'} = smbReadFile("Model-S/stk1_fw_glbp.txt");
						$anchor{'stk2_fw_int_gen'} = smbReadFile("Model-S/stk2_fw_glbp.txt");
					} else {
						$anchor{'cis1_fw_int_gen'} = '';
						$anchor{'cis2_fw_int_gen'} = '';
						$anchor{'stk1_fw_int_gen'} = '';
						$anchor{'stk2_fw_int_gen'} = '';
					}
				} else {
					if ( $anchor{'fw'} eq 'Y' ) {
						$anchor{'cis1_fw_int_gen'} = smbReadFile("Model-S/cis1_fw_glbp.txt");
						$anchor{'cis2_fw_int_gen'} = smbReadFile("Model-S/cis2_fw_glbp.txt");
						$anchor{'stk1_fw_int_gen'} = smbReadFile("Model-S/stk1_fw_glbp.txt");
						$anchor{'stk2_fw_int_gen'} = smbReadFile("Model-S/stk2_fw_glbp.txt");
					} else {
						$anchor{'cis1_fw_int_gen'} = '';
						$anchor{'cis2_fw_int_gen'} = '';
						$anchor{'stk1_fw_int_gen'} = '';
						$anchor{'stk2_fw_int_gen'} = '';
					}
				}
			$anchor{'cis1_interface_stlink'} =
			cis_stl_glbp( 1, 'S', $anchor{'ips'} );
			$anchor{'cis2_interface_stlink'} =
			cis_stl_glbp( 2, 'S', $anchor{'ips'} );
			$anchor{'stk_interface_uplink'} =
			mls_upl_glbp( $switchtype, 'S', $anchor{'ips'} );
			my ( $cisone, $cistwo );
			if (    $anchor{'router_type'} eq '3945E'
				or $anchor{'router_type'} eq '3945' )
			{
				if ( $wantype eq 'Metro_Ethernet' ) {
					$cisone = 'cis1_base_3945E_metroE.txt';
					$cistwo = 'cis2_base_3945E_metroE.txt';
				} else {
					$cisone = 'cis1_base_3945E_glbp.txt';
					$cistwo = 'cis2_base_3945E_glbp.txt';
				}
			} elsif ( $anchor{'router_type'} eq '2951' ) {
				if ( $wantype eq 'Metro_Ethernet' ) {
					$cisone = 'cis1_base_2951_metroE.txt';
					$cistwo = 'cis1_base_2951_metroE.txt';
				} else {
					$cisone = 'cis1_base_2951_glbp.txt';
					$cistwo = 'cis2_base_2951_glbp.txt';
				}
			} else {
				if ( $wantype eq 'Metro_Ethernet' ) {
					$cisone = 'cis1_base_metroE.txt';
					$cistwo = 'cis2_base_metroE.txt';
				} else {
					#for upgrade ONLY
					if($anchor{'proj_type'} eq 'upgrade'){
						$cisone = 'cis1_base_glbp_upgrade.txt';
						$cistwo = 'cis2_base_glbp_upgrade.txt';
					} else {
						$cisone = 'cis1_base_glbp.txt';
						$cistwo = 'cis2_base_glbp.txt';
					}
				}
			}
		if($anchor{'proj_type'} eq 'build'){
			#VGC uplink Config
			if ( $anchor{'vgcount'} == 1 ) {
				$anchor{'vgc_interface_uplink'} = smbReadFile("Model-S/stk1_stl_vgc.txt");
			}
			elsif ( $anchor{'vgcount'} == 2 ) {
				$anchor{'vgc_interface_uplink'} = smbReadFile("Model-S/stk2_stl_vgc2.txt");
			}
			else{
				$anchor{'vgc_interface_uplink'} = '';
			}
		}


		if($anchor{'proj_type'} eq 'build'){
				# Hard coding the current_data_vlan in for P, Q and S sites, per Deven - 2016-09-30
				$anchor{'current_data_vlan'} = '201';

				# Determine stack template based on switch type
				my %stackTemplate =
				( '1', 'stk_48p_1.txt', '2', 'stk_48p_2.txt', '3', 'stk_48p_3.txt', '4', 'stk_48p_4.txt', '5', 'stk_48p_5.txt', );
				if ( defined $stackTemplate{$switchtype} ) {
					writeTemplate( "Model-S/$stackTemplate{$switchtype}", $stack . '.txt' );
					writeTemplate( "Model-S/aruba/$stackTemplate{$switchtype}", $stack . '.txt' ) if ($anchor{'stack_vendor'} eq 'aruba');
				} else {
					prtout( "Error: Switch type '$switchtype' is not a supported switch count for a S model site" );
					xit(1);
				}
				wirelessAP();
				$anchor{'tad_hosts'} = smbReadFile("Modules/s_tad_hosts.txt");
				if ( $anchor{'vgcount'} == 1 ) {
					$anchor{'vgc1_name'} =
					'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01';

					# Per Sonia M., S sites use vlan201 addressing
					$anchor{'vgc1_address'} = $anchor{'data_subnet_1'} . '.16';
					my $of = $anchor{'vgc1_name'} . '.txt';
					writeTemplate( "Misc/vgc_204XM_MDF_mls1.txt", $of );
				}
				if ( $anchor{'vgcount'} == 2 ) {
					$anchor{'vgc1_name'} = 'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01';
					$anchor{'vgc2_name'} = 'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a02';
					$anchor{'vgc1_address'} = $anchor{'data_subnet_1'} . '.16';
					$anchor{'vgc2_address'} = $anchor{'data_subnet_1'} . '.17';
					my $of = $anchor{'vgc1_name'} . '.txt';
					writeTemplate( "Misc/vgc_204XM_MDF_mls1.txt", $of );
					my $of2 = $anchor{'vgc2_name'} . '.txt';
					writeTemplate( "Misc/vgc_204XM_MDF_mls2.txt", $of2 );
				}
		}

		} elsif ( $sitetype eq 'M' or $sitetype eq 'L' or $sitetype eq 'X' ) {

			if ( scalar(@StackList) > $stacklimit ) {
				prtout( "Too many stacks indentified for this site type. Please review input form.",
						"# stacks: " . scalar(@StackList) );
				xit(1);
			}
			my %sitemask = ( 'X', '255.255.192.0', 'L', '255.255.224.0', 'M', '255.255.240.0', );
			$anchor{'site_mask'} = $sitemask{$sitetype};
			if ( $anchor{'fw'} eq 'Y' ) {
				my $null;
				( $null, $null, my $octet3 ) = split( /\./, $anchor{'site_mask'} );
				( my $fwoctet1, my $fwoctet2, my $fwoctet3 ) =
				split( /\./, $anchor{'loop_subnet'} );
				my $fwoctet3_new = $fwoctet3 + ( ( 254 - $octet3 ) / 2 );
				my $octet3_last  = $fwoctet3 + ( ( 254 - $octet3 ) / 2 );
				$anchor{'fwi_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $fwoctet3_new;
				if ( $sitetype eq 'X' ) {
					my $octetdmz = $octet3_last - 1;
					$anchor{'dmz_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $octetdmz;
				} elsif ( $sitetype eq 'L' ) {
					my $octetdmz = $octet3_last - 2;
					$anchor{'dmz_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $octetdmz;
				} elsif ( $sitetype eq 'M' ) {
					my $octetdmz = $octet3_last - 2;
					$anchor{'dmz_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $octetdmz;
				}
			}

			if( $anchor{'proj_type'} eq 'build'){
				wirelessController();
				wirelessAP();
				prtout("Wireless Configuration Complete");
				$anchor{'vlan_naming_dynamic'}      = mls_vnd();
				$anchor{'interface_define_dynamic'} = mls_idd($sitetype);
				$anchor{'mls1_vlan_define_dynamic'} = mls_vdd( 1, $sitetype );
				$anchor{'mls2_vlan_define_dynamic'} = mls_vdd( 2, $sitetype );

				if ( $anchor{'wlan'} eq 'Y' ) {
					$anchor{'mls1_wireless_dynamic'} = mls_wdd( 1, $sitetype );
					if ( $sitetype ne 'M' ) {
						$anchor{'mls2_wireless_dynamic'} = mls_wdd( 2, $sitetype );
					}
				} else {
					$anchor{'mls1_wireless_dynamic'} = '';
					$anchor{'mls2_wireless_dynamic'} = '';
				}

				# MLS changes if IPS is involved
				$anchor{'mls1_interface_uplink'} =
				mls_upl( 1, $sitetype, $anchor{'ips'} );
				if ( $sitetype eq 'L' or $sitetype eq 'X' ) {
					$anchor{'mls2_interface_uplink'} =
					mls_upl( 2, $sitetype, $anchor{'ips'} );
				}

				#VGC uplink configs
				if ( $anchor{'vgcount'} == 1 ) {
					if ( $sitetype eq 'M'){
						$anchor{'vgc_interface_uplink'} = smbReadFile("Model-$sitetype/mls1_upl_vgc1.txt");
						}elsif( $sitetype eq 'L'){
						$anchor{'mls1_vgc_uplink'} = smbReadFile("Model-$sitetype/mls1_upl_vgc.txt");
						$anchor{'mls2_vgc_uplink'} = '!';
						}
				}
				elsif($anchor{'vgcount'} == 2 ){
					if ( $sitetype eq 'M'){
						$anchor{'vgc_interface_uplink'} = smbReadFile("Model-$sitetype/mls1_upl_vgc2.txt");
					}elsif( $sitetype eq 'L'){
						$anchor{'mls1_vgc_uplink'} = smbReadFile("Model-$sitetype/mls1_upl_vgc.txt");
						$anchor{'mls2_vgc_uplink'} = smbReadFile("Model-$sitetype/mls2_upl_vgc.txt");
					}
				}
				else{
					if ( $sitetype eq 'M'){
						$anchor{'vgc_interface_uplink'} = '!';
					}elsif( $sitetype eq 'L'){
						$anchor{'mls1_vgc_uplink'} = '!';
						$anchor{'mls2_vgc_uplink'} = '!';
					}
				}

				# MLS and Stack configs
				prtout("Writing MLS Configurations");
				my $template = "Model-$sitetype" . '/mls1_' . $sitetype . '.txt';
				writeTemplate( $template, $anchor{'mls1_name'} . '.txt' );
				if ( $sitetype eq 'L' or $sitetype eq 'X' ) {
					$template = "Model-$sitetype" . '/mls2_' . $sitetype . '.txt';
					writeTemplate( $template, $anchor{'mls2_name'} . '.txt' );
				}
			}

			##### CIS configs
			my $cisone = 'none';
			my $cistwo = 'none';


			$anchor{'tad_mls1_suffix'} = 'a01';    # L or X site types
			$anchor{'tad_mls2_suffix'} = 'a02';
			if ( $sitetype eq 'M' ) {
				$anchor{'tad_mls1_suffix'} = 'a01-1';    # M site type
				$anchor{'tad_mls2_suffix'} = 'a01-2';
			}

			#XLM TAD hosts
			$anchor{'tad_hosts'} = smbReadFile("Modules/xlm_tad_hosts.txt");

			if( $anchor{'proj_type'} eq 'build'){
				if ( $anchor{'vgcount'} > 0 ) {
					$anchor{'vgc1_name'} =
					'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01';
					my $of = $anchor{'vgc1_name'} . '.txt';
					$anchor{'vgc1_address'} = $anchor{'svr_subnet_1'} . '.16';
					writeTemplate( "Misc/vgc_204XM_MDF_mls1.txt", $of );
				}
				if ( $anchor{'vgcount'} == 2 ) {
					$anchor{'vgc2_name'} =
					'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a02';
					my $of = $anchor{'vgc2_name'} . '.txt';
					$anchor{'vgc2_address'} = $anchor{'svr_subnet_1'} . '.17';
					writeTemplate( "Misc/vgc_204XM_MDF_mls2.txt", $of );
				}
			}
		}

		if( $anchor{'proj_type'} eq 'build'){
				# TAD template (all site types use the same template)
				setTadPorts();    # set tad port names
				writeTemplate( "Generic-OOB/tad_g526.txt", 'tad' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01_G526.txt' );
				writeTemplate( "Generic-OOB/g526.txt", 'lte' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01.txt' );

			if ( $sitetype eq 'P' or $sitetype eq 'Q' ) {
				writeVisioPQ( $sitetype, 'Normal' );
				writeVisioPQ( $sitetype, 'Clean' );
			} elsif ( $sitetype eq 'S' ) {

				# The 'overtab' data needs to be carried from the normal to the clean tab
				writeVisioS( $sitetype, 'Normal' );
				writeVisioS( $sitetype, 'Clean' );
			} elsif ( $sitetype eq 'X' or $sitetype eq 'L' or $sitetype eq 'M' ) {
				my %stacklimit = ( 'M', 4, 'L', 8, 'X', 16 );    # these site types need the stack limit to create the Visio

				# For whatever reason the 'biotab' state needs to be passed to the clean version, then added to again
				writeVisioXLM( $sitetype, $stacklimit{$sitetype}, 'Normal' );
				writeVisioXLM( $sitetype, $stacklimit{$sitetype}, 'Clean' );
			}


			prtout("Writing IP Summary XLS");
			my $ipSummaryFile;
			if ( $anchor{'uhgdivision'} eq 'UHG' ) {
				$ipSummaryFile = writeIPSummary( $anchor{'site_code'}, $sitetype, $stacklimit );
				writeCISummary( $anchor{'site_code'}, $sitetype, $stacklimit );
			} else {
				$ipSummaryFile = writeTricareIPSummary( $anchor{'site_code'}, $sitetype, $stacklimit );
			}
			writeEquipmentValidation($anchor{'site_code'});
			writeRemoteSiteBuildChecklist($anchor{'site_code'});
			unlink("$ROOTDIR/Files/$ipSummaryFile");

			#writing sdwan for build
			writeSDWANcsv($anchor{'site_code'},$anchor{'cis1_name'},$anchor{'cis2_name'},$anchor{'router_seltype'}, $anchor{'int_type'}, $anchor{'int_type_r1'}, $anchor{'int_type_r2'}, $anchor{'transport'});

		}elsif($anchor{'proj_type'} eq 'proj-sdwan'){
			#writing sdwan template for proj-sdwan only
			writeSDWANcsv($anchor{'site_code'},$anchor{'cis1_name'},$anchor{'cis2_name'},$anchor{'router_seltype'}, $anchor{'int_type'}, $anchor{'int_type_r1'}, $anchor{'int_type_r2'}, $anchor{'transport'});
		}

		my $zipfile = compress();

		prtout( "Configurator Output Generation Complete.<br/>",
				"<a HREF='/tmp/$OutputDir.zip' >D&E and IP Summary can be found here</a>" );
	}

sub writeVisioPQ {    # was handlePsiteVisio, handleTriPsiteVisio and handleQsiteVisio

	return unless ($anchor{'proj_type'} eq 'build');
		$anchor{'street'} =~s/&/&amp;/g;

	( my $sitetype, my $vistype ) = @_;
	$sitetype =~ tr/A-Z/a-z/;

	# new code since P & Q site output was combined - Q site should not be TriCare
	if ( $sitetype eq 'q' and $anchor{'uhgdivision'} eq 'TRICARE' ) {
		prtout("Error; Q site should not use the TriCare division");
		return;
	}

	# Set the template and output file
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-" . $sitetype . '-template.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-" . $sitetype . '-template-clean.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';
	}

	# UHG division
	my $backbiotab   = 0;
	my $backmdftab   = 4;
	my $mdfpicotab   = 5;
	my $backipttab   = 6;
	my $biographytab = 7;

	my $cistab = 8;    # generic for P and & sites
	my $ipttab = 10;
	my $tadtab = 13;

	# Tricare
	$tadtab = 9 if ( $anchor{'uhgdivision'} eq 'TRICARE' );
	prtout("Opening Visio Template for Processing");
	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );    #zzz error handling
	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);

	# Upload tabs
	my %tabs = VisioReadTabs( \$fill );

	# UHG division
	my $mdftab =
	    $tabs{$biographytab}
	  . $tabs{$mdfpicotab}
	  . $tabs{$tadtab}
	  . $tabs{$ipttab}
	  . $tabs{$cistab}
	  . $tabs{$backbiotab}
	  . $tabs{$backmdftab}
	  . $tabs{$backipttab};

	# Tricare
	if ( $anchor{'uhgdivision'} eq 'TRICARE' ) {
		$mdftab =
		    $tabs{$biographytab}
		  . $tabs{$mdfpicotab}
		  . $tabs{$cistab}
		  . $tabs{$tadtab}
		  . $tabs{$ipttab}
		  . $tabs{$backbiotab}
		  . $tabs{$backmdftab}
		  . $tabs{$backipttab};
	}
	VisioControlLayer( 'WLAN', 0, \$mdftab ) if ( $anchor{'wlan'} ne 'Y' );

	# Search and replace variables
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
	}

	# Logic for wae on MDF tab (UHG division only)
	VisioControlLayer( 'wae', 0, \$mdftab ) # Removing WAAS per Ed.
	  if ( $anchor{'uhgdivision'} eq 'UHG' );

	#VGC drawings
	VisioControlLayer( 'VGC', 0, \$mdftab ); 	#set vgc default to 0
	if ( $anchor{'vgcount'} > 0 ){
		VisioControlLayer( 'VGC', 1, \$mdftab );
	}

	# Put together tabs
	my $generatedTabs = $mdftab;
	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified $vistype Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");
	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );             #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

 #	`$SMBCLIENT -U $SMB_USER%$SMB_PASS -W MS //unpiox56pn/netsvcs/ -c 'put $ROOTDIR/Files/$vOutput $SMB_FIN_DIR/$OutputDir/$vOutput`;
	unlink("$ROOTDIR/Files/$vOutput");
}

sub writeVisioS {                                                 # was handleSsiteVisio
return unless ($anchor{'proj_type'} eq 'build');

	$anchor{'street'} =~s/&/&amp;/g;




	( my $sitetype, my $vistype, my $overtab ) = @_;
	$overtab = '' if ( !( defined $overtab ) );
	my $ips  = $anchor{'ips'};
	my $fw   = $anchor{'fw'};
	my $wlan = $anchor{'wlan'};

	# Set template and output file
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-s-template.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-s-template-clean.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';
	}
	my $backmdftab   = 4;
	my $mdfsmalltab  = 5;
	my $backipttab   = 6;
	my $backbiotab   = 9;
	my $biographytab = 10;
	my $tadtab       = 14;
	my $fwoverinttab = 15;
	my $backsectab   = 16;
	my $secl3tab     = 17;
	my $stkovertab   = 19;
	my $waetab       = 22;
	my $ipttab       = 24;
	my $fwoverexttab = 26;
	my $routertab    = 23;    # New 4/20/2017 - replaces router specific tabs
	my $cis2951tab   = 8;     # replaced by $routertab
	my $cis4451tab   = 8;     # replaced by $routertab
	my $cis3945tab   = 23;    # replaced by $routertab
	prtout("Opening Visio Template for Processing");
	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );    #zzz error handling

	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);

	# Upload tabs
	my %tabs = VisioReadTabs( \$fill );

	# Define tabs to variables
	my $biotab = $tabs{$biographytab};

	# Add TAD tab for sites with IPS
	my $mdftab = $tabs{$mdfsmalltab} . $tabs{$tadtab};

	# Security Layer3 and FW Overviews
	$overtab .= $tabs{$secl3tab} . $tabs{$fwoverinttab} if ( $fw eq 'Y' );

	# WAAS overview, IPT Templat eand Stack overview
	$overtab .= $tabs{$waetab} . $tabs{$ipttab} . $tabs{$stkovertab} ;

	# 4/20/2017 - simplified code
	#           - also using $routertab instead of router specific tabs
	VisioControlLayer( '2951',                 0, \$mdftab );
	VisioControlLayer( '3945',                 0, \$mdftab );
	VisioControlLayer( '4321',                 0, \$mdftab );
	VisioControlLayer( '4331',                 0, \$mdftab );
	VisioControlLayer( '4351',                 0, \$mdftab );
	VisioControlLayer( '4451',                 0, \$mdftab );
	VisioControlLayer( 'C8200-1N-4T',          0, \$mdftab );
	VisioControlLayer( 'SDWAN',                0, \$mdftab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$mdftab );
	$overtab .= $tabs{$routertab};
	VisioControlLayer( '2951',                 0, \$overtab );
	VisioControlLayer( '3945',                 0, \$overtab );
	VisioControlLayer( '3945E',                0, \$overtab );
	VisioControlLayer( '4321',                 0, \$overtab );
	VisioControlLayer( '4331',                 0, \$overtab );
	VisioControlLayer( '4351',                 0, \$overtab );
	VisioControlLayer( '4451',                 0, \$overtab );
	VisioControlLayer( 'C8200-1N-4T',          0, \$overtab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$overtab );

	VisioControlLayer( 'aruba',                 0, \$overtab );
	VisioControlLayer( 'cisco',                 0, \$overtab );
	VisioControlLayer( $anchor{'stack_vendor'}, 1, \$overtab );

	my $switchtype = ( split( /,/, $StackList[0] ) )[1];
	VisioControlLayer( 'stack1', 0, \$mdftab );
	VisioControlLayer( 'stack2', 0, \$mdftab );
	VisioControlLayer( 'stack3', 0, \$mdftab );
	VisioControlLayer( 'stack4', 0, \$mdftab );
	VisioControlLayer( 'stack5', 0, \$mdftab );
	VisioControlLayer( 'aruba_stack1', 0, \$mdftab );
	VisioControlLayer( 'aruba_stack2', 0, \$mdftab );
	VisioControlLayer( 'aruba_stack3', 0, \$mdftab );
	VisioControlLayer( 'aruba_stack4', 0, \$mdftab );
	VisioControlLayer( 'aruba_stack5', 0, \$mdftab );
	VisioControlLayer( 'aruba_stack1_port', 0, \$mdftab );
	VisioControlLayer( 'stack1_port', 0, \$mdftab );
	VisioControlLayer( 'vgc2_port', 0, \$mdftab );
	VisioControlLayer( 'vgc2_port_aruba', 0, \$mdftab );

	# Logic for switch layers
	if ( $switchtype eq '24-port' ) {
		VisioControlLayer( 'stack2_cable',    0, \$mdftab );
		VisioControlLayer( 'stack3_cable',    0, \$mdftab );
		VisioControlLayer( 'stack4_cable',    0, \$mdftab );
		VisioControlLayer( 'stack5_cable',    0, \$mdftab );
		VisioControlLayer( 'stack1_cable_fw', 0, \$mdftab ) if ( $fw ne 'Y' );
		VisioControlLayer( 'stack2_cable_fw', 0, \$mdftab );
		if ($anchor{'stack_vendor'} eq 'aruba' ){
				VisioControlLayer( 'aruba_stack1', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack1_port', 1, \$mdftab );
		}
		else {
				VisioControlLayer( 'stack1',          1, \$mdftab );
				VisioControlLayer( 'stack1_port', 1, \$mdftab );
				VisioControlLayer( 'aruba', 0, \$mdftab );

		}
	} elsif ( $switchtype eq '1' ) {
		VisioControlLayer( 'stack2_cable',    0, \$mdftab );
		VisioControlLayer( 'stack3_cable',    0, \$mdftab );
		VisioControlLayer( 'stack4_cable',    0, \$mdftab );
		VisioControlLayer( 'stack5_cable',    0, \$mdftab );
		VisioControlLayer( 'stack24port',     0, \$mdftab );
		VisioControlLayer( 'wae_stk1',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'wae_stk2',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'stack1_cable_fw', 0, \$mdftab ) if ( $fw ne 'Y' );
		VisioControlLayer( 'stack2_cable_fw', 0, \$mdftab );
		if ($anchor{'stack_vendor'} eq 'aruba' ){
				VisioControlLayer( 'aruba_stack1', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack1_port', 1, \$mdftab );
		}
		else {
				VisioControlLayer( 'stack1',          1, \$mdftab );
				VisioControlLayer( 'stack1_port', 1, \$mdftab );
				VisioControlLayer( 'aruba', 0, \$mdftab );
		}
	} elsif ( $switchtype eq '2' ) {
		VisioControlLayer( 'stack1_cable',    0, \$mdftab );
		VisioControlLayer( 'stack3_cable',    0, \$mdftab );
		VisioControlLayer( 'stack4_cable',    0, \$mdftab );
		VisioControlLayer( 'stack5_cable',    0, \$mdftab );
		VisioControlLayer( 'stack24port',     0, \$mdftab );
		VisioControlLayer( 'wae_stk1',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'wae_stk2',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'stack1_cable_fw', 0, \$mdftab );
		VisioControlLayer( 'stack2_cable_fw', 0, \$mdftab ) if ( $fw ne 'Y' );
		if ($anchor{'stack_vendor'} eq 'aruba' ){
				VisioControlLayer( 'aruba_stack1', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack2', 	  1, \$mdftab ) ;
		}
		else {
				VisioControlLayer( 'stack1',          1, \$mdftab );
				VisioControlLayer( 'stack2',          1, \$mdftab );
				VisioControlLayer( 'aruba', 0, \$mdftab );
		}
	} elsif ( $switchtype eq '3' ) {
		VisioControlLayer( 'stack1_cable',    0, \$mdftab );
		VisioControlLayer( 'stack4_cable',    0, \$mdftab );
		VisioControlLayer( 'stack5_cable',    0, \$mdftab );
		VisioControlLayer( 'stack24port',     0, \$mdftab );
		VisioControlLayer( 'wae_stk1',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'wae_stk2',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'stack1_cable_fw', 0, \$mdftab );
		VisioControlLayer( 'stack2_cable_fw', 0, \$mdftab ) if ( $fw ne 'Y' );
		if ($anchor{'stack_vendor'} eq 'aruba' ){
				VisioControlLayer( 'aruba_stack1', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack2', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack3', 	  1, \$mdftab ) ;
		}
		else {
				VisioControlLayer( 'stack1',          1, \$mdftab );
				VisioControlLayer( 'stack2',          1, \$mdftab );
				VisioControlLayer( 'stack3',          1, \$mdftab );
				VisioControlLayer( 'aruba', 0, \$mdftab );
		}
	} elsif ( $switchtype eq '4' ) {
		VisioControlLayer( 'stack1_cable',    0, \$mdftab );
		VisioControlLayer( 'stack5_cable',    0, \$mdftab );
		VisioControlLayer( 'stack24port',     0, \$mdftab );
		VisioControlLayer( 'wae_stk1',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'wae_stk2',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'stack1_cable_fw', 0, \$mdftab );
		VisioControlLayer( 'stack2_cable_fw', 0, \$mdftab ) if ( $fw ne 'Y' );
		if ($anchor{'stack_vendor'} eq 'aruba' ){
				VisioControlLayer( 'aruba_stack1', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack2', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack3', 	  1, \$mdftab ) ;
				VisioControlLayer( 'aruba_stack4', 	  1, \$mdftab ) ;
		}
		else {
				VisioControlLayer( 'stack1',          1, \$mdftab );
				VisioControlLayer( 'stack2',          1, \$mdftab );
				VisioControlLayer( 'stack3',          1, \$mdftab );
				VisioControlLayer( 'stack4',          1, \$mdftab );
				VisioControlLayer( 'aruba', 0, \$mdftab );
		}
	} elsif ( $switchtype eq '5' ) {
		VisioControlLayer( 'stack1_cable',    0, \$mdftab );
		VisioControlLayer( 'stack24port',     0, \$mdftab );
		VisioControlLayer( 'wae_stk1',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'wae_stk2',        0, \$mdftab );                     # Additional layer for was02's stk prot
		VisioControlLayer( 'stack1_cable_fw', 0, \$mdftab );
		VisioControlLayer( 'stack2_cable_fw', 0, \$mdftab ) if ( $fw ne 'Y' );
	}
	VisioControlLayer( 'wae', 0, \$mdftab ); #removing WAAS per ed
	my $backtab = $tabs{$backmdftab} . $tabs{$backipttab} . $tabs{$backbiotab} . $tabs{$backsectab};

	VisioControlLayer( 'VGC',  0, \$mdftab ); 	#VGC drawing, set vgc default to 0
	VisioControlLayer( 'VGC2', 0, \$mdftab ); 	#VGC2 drawing, set vgc default to 0
	VisioControlLayer( 'vgc2_port_aruba', 0, \$mdftab );
	VisioControlLayer( 'vgc2_port', 0, \$mdftab );
	VisioControlLayer( 'vgc_port', 0, \$mdftab );
	VisioControlLayer( 'vgc_port_aruba', 0, \$mdftab );

	if ( $anchor{'vgcount'} == 1 ){
	VisioControlLayer( 'VGC', 1, \$mdftab );
	VisioControlLayer( 'vgc_port', 1, \$mdftab ) if ($anchor{'stack_vendor'} eq 'cisco');
	VisioControlLayer( 'vgc_port_aruba', 1, \$mdftab ) if ($anchor{'stack_vendor'} eq 'aruba');
	}
	elsif ( $anchor{'vgcount'} == 2 ){
	VisioControlLayer( 'VGC',  1, \$mdftab );
	VisioControlLayer( 'VGC2', 1, \$mdftab );
	VisioControlLayer( 'vgc_port', 1, \$mdftab ) if ($anchor{'stack_vendor'} eq 'cisco');
	VisioControlLayer( 'vgc2_port', 1, \$mdftab )if ($anchor{'stack_vendor'} eq 'cisco');
	VisioControlLayer( 'vgc_port_aruba', 1, \$mdftab ) if ($anchor{'stack_vendor'} eq 'aruba');
	VisioControlLayer( 'vgc2_port_aruba', 1, \$mdftab ) if ($anchor{'stack_vendor'} eq 'aruba');
	}

# SDWAN
	VisioControlLayer( '2_transports', 0, \$mdftab )
	if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 );

	VisioControlLayer( '3_transports', 0, \$mdftab )
	if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 );

	VisioControlLayer( 'non_sdwan', 0, \$mdftab ) if ($anchor{'SDWAN'} eq 'Y');

	# Keep or remove IPS layer
	VisioControlLayer( 'IPS',   0, \$mdftab ) if ( $ips ne 'Y' );
	VisioControlLayer( 'NOIPS', 0, \$mdftab ) if ( $ips ne 'N' );
	VisioControlLayer( 'FW',    0, \$mdftab ) if ( $fw ne 'Y' );
	VisioControlLayer( 'WLAN',  0, \$mdftab ) if ( $wlan ne 'Y' );

	# Search and replace variables
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
		$overtab =~ s/\!$key\!/$anchor{$key}/g;
		$backtab =~ s/\!$key\!/$anchor{$key}/g;
	}

	# Put together tabs
	my $generatedTabs = $biotab . $mdftab . $overtab . $backtab;
	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified $vistype Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");

	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );             #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

	unlink("$ROOTDIR/Files/$vOutput");
	return $overtab;
}

sub writeVisioXLM {                                               # was handleXLMsiteVisio
	return unless ($anchor{'proj_type'} eq 'build');
		$anchor{'street'} =~s/^ | &//g;
		$anchor{'street'} =~s/&/&amp;/g;

	( my $sitetype, my $stacklimit, my $vistype ) = @_;
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-xlm-template.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-xlm-template-clean.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';
	}
	my $idf2tab       = 0;
	my $idf3tab       = 4;
	my $backmdftab    = 5;
	my $mdfnextab     = 6;
	my $mdflargetab   = 7;
	my $idf1tab       = 8;
	my $mdfmediumtab  = 9;
	my $backstktab    = 10;
	my $backidfxab    = 11;
	my $mdfstktab     = 12;
	my $backipttab    = 14;
	my $backbiotab    = 16;
	my $biographytab  = 17;
	my $cis3945tab    = 19;
	my $backsectab    = 24;
	my $tadtab        = 27;
	my $backwlantab   = 30;
	my $wirelesstab   = 31;
	my $stkovertab    = 33;
	my $waetab        = 34;
	my $ipttab        = 36;
	my $routerovertab = 45;
	my $fwovertab     = 996;
	my $dmzflowtab    = 997;
	my $untrusttab    = 998;
	my $secl3tab      = 999;

	my @idfxabs = ();
	my $thisidf = '';
	my $lastidf = '';
	my $store   = '';

	# Save stack names (space separated) for each unique idf. I think. The old code does this in a weird way
	# and I might have missed something.
	my @idforder;
	my $stacklist = '';
	foreach my $stk ( @StackList, 'end' ) {    # Adding 'end' forces an update in the final iteration of the loop.
		                                       # 'end' can be anything other than a valid stack name.
		my $idf = substr( $stk, 9, 3 );
		if ( $lastidf eq '' ) {                # first stack in loop iteration
			$lastidf   = $idf;
			$stacklist = $stk;
		} elsif ( $idf ne $lastidf ) {         # When the stack name has changed,
			push @idforder, $stacklist;        # save off the stack list
			$stacklist = $stk;                 # and reset the stack list to the current stack name.
		} else {
			$stacklist .= " $stk";             # Adds stack name to existing (space delimited) list
		}
		$lastidf = $idf;
	}
	my $idfxabid = 50;
	my $idfct    = 0;
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $stack    = $ct + 1;
		my $name     = 'sc' . $stack;
		my $stkcount = ( split( /,/, $StackList[$ct] ) )[1];
		$anchor{$name} = $stkcount;
	}
	my $ips  = $anchor{'ips'};
	my $fw   = $anchor{'fw'};
	my $ipt  = $anchor{'ipt'};
	my $wlan = $anchor{'wlan'};
	prtout("Opening Visio Template for Processing");
	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );

	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);
	my %tabs = VisioReadTabs( \$fill );

	# Add Bio Page
	my $biotab .= $tabs{$biographytab};

	# Add MDF tabs
	# show stack overview sitetype layer
	$mdfstktab = $tabs{$mdfstktab};
	VisioControlLayer( 'idfx', 0, \$mdfstktab ) if ( $sitetype ne 'X' );
	VisioControlLayer( 'idfl', 0, \$mdfstktab ) if ( $sitetype ne 'L' );
	VisioControlLayer( 'idfm', 0, \$mdfstktab ) if ( $sitetype ne 'M' );

	my $mdftab;
	if ( $sitetype eq 'M' ) {
		$mdftab = $tabs{$mdfmediumtab} . $mdfstktab;
	} elsif ( $sitetype eq 'L' or $sitetype eq 'X' ) {
		$mdftab = $tabs{$mdflargetab} . $mdfstktab;
	}

	# Set all to hidden by default
	VisioControlLayer( '2951',  0, \$mdftab );
	VisioControlLayer( '3945',  0, \$mdftab );
	VisioControlLayer( '3945E', 0, \$mdftab );
	VisioControlLayer( '4321',  0, \$mdftab );
	VisioControlLayer( '4331',  0, \$mdftab );
	VisioControlLayer( '4351',  0, \$mdftab );
	VisioControlLayer( '4451',  0, \$mdftab );
	VisioControlLayer( '4461',  0, \$mdftab );
	VisioControlLayer( 'ASR',   0, \$mdftab );
	VisioControlLayer( 'C8200-1N-4T',  0, \$mdftab );
	VisioControlLayer( 'C8300-1N1S',   0, \$mdftab );
	VisioControlLayer( 'C8300-2N2S',   0, \$mdftab );
	VisioControlLayer( 'VGC1',  0, \$mdftab );
	VisioControlLayer( 'VGC2',  0, \$mdftab );
	VisioControlLayer( 'SDWAN', 0, \$mdftab );

	# Make the appropriate layer visible
	VisioControlLayer( $anchor{'router_type'}, 1, \$mdftab );
	if ( $sitetype eq 'L' ) {
		VisioControlLayer( 'mdfl', 1, \$mdftab );
		VisioControlLayer( 'mdfx', 0, \$mdftab );
	}
	#VGC drawing
	if ( $anchor{'vgcount'} == 1 ){
		VisioControlLayer( 'VGC1', 1, \$mdftab );
	}
	elsif( $anchor{'vgcount'} == 2 ){
		VisioControlLayer( 'VGC1', 1, \$mdftab );
		VisioControlLayer( 'VGC2', 1, \$mdftab );
	}

	##SDWAN
		VisioControlLayer( '2_transports', 0, \$mdftab )
		if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 );

		VisioControlLayer( '3_transports', 0, \$mdftab )
		if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 );

		if ($anchor{'SDWAN'} eq 'Y') {
			VisioControlLayer( 'non_SDWAN_con',	 	0, \$mdftab ); #disable non sdwan connectors
			VisioControlLayer( 'SDWAN', 			1, \$mdftab ); #enable the SDWAN router
		}
		if ($anchor{'SDWAN'} eq 'Y' and $sitetype =~ /^[XL]$/) {
			VisioControlLayer( 'SDWAN', 			1, \$mdftab ); #enable the SDWAN router
		}

	VisioControlLayer( 'wae',  0, \$mdftab ); #removing WAAS per ed
	VisioControlLayer( 'idfx', 0, \$mdftab ) if ( $sitetype ne 'X' );
	VisioControlLayer( 'idfl', 0, \$mdftab ) if ( $sitetype ne 'L' );
	VisioControlLayer( 'idfm', 0, \$mdftab ) if ( $sitetype ne 'M' );

	# Layer control for IPS and IPT
	VisioControlLayer( 'IPS',   0, \$mdftab ) if ( $ips ne 'Y' );
	VisioControlLayer( 'NOIPS', 0, \$mdftab ) if ( $ips ne 'N' );
	VisioControlLayer( 'FW',    0, \$mdftab ) if ( $fw ne 'Y' );
	VisioControlLayer( 'WLAN',  0, \$mdftab ) if ( $wlan ne 'Y' );

	# Layer control for stacks on MDF tab - configurator allows 22 stacks maximum
	my $stkct = scalar(@StackList);
	for ( my $ct = 22 ; $ct > 1 ; $ct-- ) {
		VisioControlLayer( 'mdf_stack' . $ct, 0, \$mdftab ) if ( $stkct < $ct );
	}

	# Search and replace values in mdf and bio tabs
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
		$biotab =~ s/\!$key\!/$anchor{$key}/g;
	}
	my $generatedTabs = $biotab . $mdftab . $tabs{$ipttab};
	my ( $swcount, $workingtab );
	foreach my $idf (@idforder) {
		my @switches = split( /\s/, $idf );
		$swcount    = scalar(@switches);
		$workingtab = $tabs{$idf1tab} if ( $swcount == 1 );    # one stack per IDF
		$workingtab = $tabs{$idf2tab} if ( $swcount == 2 );    # one stack per IDF
		$workingtab = $tabs{$idf3tab} if ( $swcount == 3 );    # one stack per IDF
		if ( scalar(@switches) > 3 ) {                         # limited to 3 to keep Visio page clean looking
			prtout( "There are too many stacks in this IDF for Configurator to handle.",
					"Please split them into two IDFs and retrun." );
			xit(1);
		}
		my $idfname =
		  substr( $switches[0], 8, 3 ) . '-' . substr( $switches[0], 11, 1 );
		VisioRenameTab( "IDF " . $idfname, \$workingtab );
		$idfxabid++;
		VisioReIDTab( $idfxabid, \$workingtab );

		#zzz copied (more or less) from old code - find a cleaner way to do this
		for ( my $peridf = 2 ; $peridf >= 0 ; $peridf-- ) {
			if ( $swcount > $peridf ) {
				my $idfct = $idfct + $peridf + 1;
				my $val2  = $peridf + 1;
				my $key   = 'idf' . $idfct . '_ds_1';
				my $val   = $anchor{$key};
				my $rval  = "!idf" . $val2 . '_ds_1!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_es_1';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '_es_1!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_vs_1';
				$val  = $anchor{$key};
				$rval = "!idf" . $val2 . '_vs_1!';
				$workingtab =~ s/$rval/$val/g;
				my ( $stkname, $swcount ) = split( /,/, $switches[$peridf] );
				$rval = "!stack" . $val2 . '_name!';
				$workingtab =~ s/$rval/$stkname/g;
				doswitches( $swcount, $val2, \$workingtab );
				$key  = 'idf' . $idfct . '_Data_vlan_1';
				$val  = $anchor{$key};
				$rval = '!dvlan' . $val2 . '!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_EAC_vlan_1';
				$val  = $anchor{$key};
				$rval = '!evlan' . $val2 . '!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_voice_vlan_1';
				$val  = $anchor{$key};
				$rval = "!vvlan" . $val2 . '!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_dlo_mls1';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '_dlo1!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_dlo_mls2';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '_dlo2!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_elo_mls1';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '_elo1!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_elo_mls2';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '_elo2!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_vlo_mls1';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '_vlo1!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_vlo_mls2';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '_vlo2!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '_mls_up';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . 'up!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '-1_mls_up';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '-1up!';
				$workingtab =~ s/$rval/$val/g;
				$key  = 'idf' . $idfct . '-2_mls_up';
				$val  = $anchor{$key};
				$rval = '!idf' . $val2 . '-2up!';
				$workingtab =~ s/$rval/$val/g;
			}
		}
		$workingtab =~ s/\!idf_name\!/$idfname/g;
		foreach my $key ( keys %anchor ) {
			$workingtab =~ s/\!$key\!/$anchor{$key}/g;
		}
		VisioControlLayer( 'idfx', 0, \$workingtab ) if ( $sitetype ne 'X' );
		VisioControlLayer( 'idfl', 0, \$workingtab ) if ( $sitetype ne 'L' );
		VisioControlLayer( 'idfm', 0, \$workingtab ) if ( $sitetype ne 'M' );
		$generatedTabs .= $workingtab;
		$idfct += $swcount;
	}

	# Add background tabs
	my $idfbg = $tabs{$backidfxab};

	# Disable unnecessary IDF layers
	VisioControlLayer( 'idfx', 0, \$idfbg ) if ( $sitetype ne 'X' );
	VisioControlLayer( 'idfl', 0, \$idfbg ) if ( $sitetype ne 'L' );
	VisioControlLayer( 'idfm', 0, \$idfbg ) if ( $sitetype ne 'M' );
	my $mdfbg  = $tabs{$backmdftab};
	my $stabg  = $tabs{$backstktab};
	my $iptbg  = $tabs{$backipttab};
	my $biobg  = $tabs{$backbiotab};
	my $secbg  = $tabs{$backsectab};
	my $wlanbg = $tabs{$backwlantab};
	if ($anchor{'wlc_model'} ne '9800'){
		#Make selection for 'idfxl_old'
		if ($sitetype =~ /^(X|L)$/){
			VisioControlLayer( 'idfxl', 0, \$wlanbg );
			VisioControlLayer( 'idfm', 0, \$wlanbg );
			VisioControlLayer( 'idfm_2', 0, \$wlanbg );
			VisioControlLayer( 'idfm_old', 0, \$wlanbg );
		}
		#Make selection for 'idfm_old'
		else{
			VisioControlLayer( 'idfxl', 0, \$wlanbg );
			VisioControlLayer( 'idfm', 0, \$wlanbg );
			VisioControlLayer( 'idfm_2', 0, \$wlanbg );
			VisioControlLayer( 'idfxl_old', 0, \$wlanbg );
		}
	}elsif ($sitetype =~ /^M$/){
		if ( ($anchor{'wlc_nmbr'} == 1) ){
		VisioControlLayer( 'idfm_2', 0, \$wlanbg );
		VisioControlLayer( 'idfxl', 0, \$wlanbg );
		VisioControlLayer( 'idfxl_old', 0, \$wlanbg );
		VisioControlLayer( 'idfm_old', 0, \$wlanbg );
		}elsif ( ($anchor{'wlc_nmbr'} == 2) ){
		VisioControlLayer( 'idfm', 0, \$wlanbg );
		VisioControlLayer( 'idfxl', 0, \$wlanbg );
		VisioControlLayer( 'idfxl_old', 0, \$wlanbg );
		VisioControlLayer( 'idfm_old', 0, \$wlanbg );
		}
	}else{
	#VisioControlLayer( 'idftg', 0, \$wlanbg ) if ( $sitetype !~ /^(T|G)$/ );
	VisioControlLayer( 'idfxl_old', 0, \$wlanbg );
	VisioControlLayer( 'idfm_old', 0, \$wlanbg );
	VisioControlLayer( 'idfm', 0, \$wlanbg );
	VisioControlLayer( 'idfm_2', 0, \$wlanbg );
	}

	#Visio tad tabs
	my $tad = $tabs{$tadtab};
	VisioControlLayer( 'idfx', 0, \$tad ) if ( $sitetype ne 'X' );
	VisioControlLayer( 'idfl', 0, \$tad ) if ( $sitetype ne 'L' );
	VisioControlLayer( 'idfm', 0, \$tad ) if ( $sitetype ne 'M' );

	# Do substitutions for these tabs
	foreach my $key ( keys %anchor ) {
		$idfbg =~ s/\!$key\!/$anchor{$key}/g;
		$mdfbg =~ s/\!$key\!/$anchor{$key}/g;
		$stabg =~ s/\!$key\!/$anchor{$key}/g;
		$iptbg =~ s/\!$key\!/$anchor{$key}/g;
		$biobg =~ s/\!$key\!/$anchor{$key}/g;
		$secbg =~ s/\!$key\!/$anchor{$key}/g;
		$wlanbg =~ s/\!$key\!/$anchor{$key}/g;
		$tad =~ s/\!$key\!/$anchor{$key}/g;
	}
	$generatedTabs .= $tad;

	# Wireless tab
	if ( $wlan eq 'Y' ) {
		my $wlantab = $tabs{$wirelesstab};
		foreach my $key ( keys %anchor ) {
			$wlantab =~ s/\!$key\!/$anchor{$key}/g;
		}
		if ( $anchor{'wlc_model'} ne '9800' ){
			VisioControlLayer( 'idfxl', 0, \$wlantab );
			VisioControlLayer( 'idfm', 0, \$wlantab );
			VisioControlLayer( 'idfm_2', 0, \$wlantab );
		}elsif ($sitetype =~ /^(M)$/){
			if ( ($anchor{'wlc_nmbr'} == 1) ){
				VisioControlLayer( 'idfm_2', 0, \$wlantab );
				VisioControlLayer( 'idfxl', 0, \$wlantab );
				VisioControlLayer( 'idfxlm_old', 0, \$wlantab );
				}elsif ( ($anchor{'wlc_nmbr'} == 2) ){
				VisioControlLayer( 'idfm', 0, \$wlantab );
				VisioControlLayer( 'idfxl', 0, \$wlantab );
				VisioControlLayer( 'idfxlm_old', 0, \$wlantab );
				}
		}else{
		VisioControlLayer( 'idfxlm_old', 0, \$wlantab );
		VisioControlLayer( 'idfm', 0, \$wlantab );
		VisioControlLayer( 'idfm_2', 0, \$wlantab );
		}

		$generatedTabs .= $wlantab;
	}
	if ( $fw eq 'Y' ) {
		my $secl3 = $tabs{$secl3tab};
		VisioControlLayer( 'idfx', 0, \$secl3 ) if ( $sitetype ne 'X' );
		VisioControlLayer( 'idfl', 0, \$secl3 ) if ( $sitetype ne 'L' );
		VisioControlLayer( 'idfm', 0, \$secl3 ) if ( $sitetype ne 'M' );
		my $fwovint = $tabs{$fwovertab};
		my $dmzflow = $tabs{$dmzflowtab};
		my $utdmz   = $tabs{$untrusttab};
		VisioControlLayer( 'idfx', 0, \$utdmz ) if ( $sitetype ne 'X' );
		VisioControlLayer( 'idfl', 0, \$utdmz ) if ( $sitetype ne 'L' );
		VisioControlLayer( 'idfm', 0, \$utdmz ) if ( $sitetype ne 'M' );

		foreach my $key ( keys %anchor ) {
			$secl3 =~ s/\!$key\!/$anchor{$key}/g;
			$fwovint =~ s/\!$key\!/$anchor{$key}/g;
			$utdmz =~ s/\!$key\!/$anchor{$key}/g;
			$dmzflow =~ s/\!$key\!/$anchor{$key}/g;
		}
		$generatedTabs .= $secl3 . $utdmz . $fwovint . $dmzflow;
	}

	#stack overview tab
	my $stacktab = $tabs{$stkovertab};
	#$overtab .= $tabs{$stkovertab};
	VisioControlLayer( 'aruba',                 0, \$stacktab );
	VisioControlLayer( 'cisco',                 0, \$stacktab );
	VisioControlLayer( $anchor{'stack_vendor'}, 1, \$stacktab );

	$routerovertab = $tabs{$routerovertab};
	VisioControlLayer( '2951',                 0, \$routerovertab );    # Set layers to hidden by default
	VisioControlLayer( '3945',                 0, \$routerovertab );
	VisioControlLayer( '3945E',                0, \$routerovertab );
	VisioControlLayer( '4321',                 0, \$routerovertab );
	VisioControlLayer( '4351',                 0, \$routerovertab );
	VisioControlLayer( '4451',                 0, \$routerovertab );
	VisioControlLayer( '4461',                 0, \$routerovertab );
	VisioControlLayer( 'ASR',                  0, \$routerovertab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$routerovertab );    # Make the appropriate layer visible
	$generatedTabs .= $routerovertab;

	$generatedTabs .= $mdfbg . $stabg . $idfbg . $iptbg . $biobg . $secbg . $wlanbg . $stacktab;
	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");

	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );             #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

 # Return this - after the 'Normal' Visio runs the 'Clean' needs to add to this value
	return $biotab;
}

sub doswitches {
return unless ($anchor{'proj_type'} eq 'build');
	( my $swct, my $chstack, my $data ) = @_;
	for ( my $ct = 5 ; $ct > 1 ; $ct-- ) {
		if ( $swct < $ct ) {
			VisioControlLayer( "idf_$chstack" . "stack_switch$ct", 0, $data );
			VisioControlLayer( "idf_$chstack" . "stack_cable$ct",  0, $data );
		}
	}
	VisioControlLayer( "idf_$chstack" . 'stack_fiber1', 0, $data )
	if ( $swct > 1 );
	VisioControlLayer( "idf_$chstack" . 'stack_cable2', 0, $data )
	  if ( $swct > 2 );
	VisioControlLayer( "idf_$chstack" . 'stack_cable3', 0, $data )
	  if ( $swct > 3 );
	VisioControlLayer( "idf_$chstack" . 'stack_cable4', 0, $data )
	  if ( $swct > 4 );

}

sub setTadPorts {
	return unless ($anchor{'proj_type'} eq 'build');

	#only for Medium, Large and XL sites
	#  TAD Ports 1-16
	$anchor{'tad_port_1'}  = $anchor{'cis1_name'};
	$anchor{'tad_port_2'}  = $anchor{'cis2_name'};
	$anchor{'tad_port_3'}  = $anchor{'mls1_name'};
	if ($sitetype eq 'M' ) {
		$anchor{'tad_port_4'}  = $anchor{'mls1_name'} . '-2';
		$anchor{'tad_int'} = 'Gi1/0/14'
	}
	if (($sitetype eq 'L' ) or ($sitetype eq 'X' )){
		$anchor{'tad_port_4'}  = $anchor{'mls2_name'};
		$anchor{'tad_int'} = 'Gi4/0/14'
	}
	$anchor{'tad_port_5'}  = $anchor{'wlc1_name'};
	$anchor{'tad_port_6'}  = $anchor{'spare_name1'};
	$anchor{'tad_port_7'}  = $anchor{'spare_name2'};
	$anchor{'tad_port_8'}  = 'Optional_8';
	$anchor{'tad_port_8'}  = 'FW1' if ( $anchor{'fw'} eq 'Y' );
	$anchor{'tad_port_9'}  = 'Optional_9';
	$anchor{'tad_port_9'}  = 'FW2' if ( $anchor{'fw'} eq 'Y' );
	$anchor{'tad_port_10'} = 'Optional_10';
	$anchor{'tad_port_10'} = $anchor{'wlc2_name'} if ( ($anchor{'wlan'} eq 'Y') and ($anchor{'wlc_nmbr'} == 2) );
	$anchor{'tad_port_11'} = 'Optional';
	$anchor{'tad_port_12'} = 'Optional';
	$anchor{'tad_port_13'} = 'Optional';
	$anchor{'tad_port_14'} = 'Optional';
	$anchor{'tad_port_15'} = 'Optional';
	$anchor{'tad_port_16'} = 'Optional';
}

sub writeIPSummary {
	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid, my $sitetype, my $stacklimit ) = @_;
	my $outputFile = $siteid . '-IP-Summary-Chart.xls';

	my $workbook = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile")
	  or die "create XLS file '$SVR_ROOTDIR/$OutputDir/$outputFile' failed: $!";

	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75);
	$ws->set_column( 'A:A', 14.5 );
	$ws->set_column( 'E:E', 16.5 );
	$ws->set_column( 'F:F', 16.5 );
	$ws->set_column( 'G:G', 16.5 );
	$ws->set_column( 'H:H', 32 );
	$ws->set_column( 'I:I', 23 );
	$ws->set_column( 'M:M', 10 );
	my $fmtBlank = $workbook->add_format( size => 12, bold => 1 );
	my $fmtGray = $workbook->add_format(
										 size      => 9,
										 bold      => 1,
										 bg_color  => 22,
										 text_wrap => 1
	);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtYellow = $workbook->add_format( bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );

	# First three rows are the header
	my $rows   = 2;
	my $height = 12;
	$ws->merge_range( 'A1:O2', "$siteid - IP Summary Chart", $fmtBlank );
	my @header = (
				   'Status',
				   'Existing User Port Count',
				   'New # of 48 port switches',
				   'New User Port Count',
				   'IDF',
				   'IP Subnet',
				   'IP Address',
				   'Device Name',
				   'Device Description',
				   'DHCP',
				   'DNS',
				   'C.Wrks',
				   'HPOpnv.',
				   'Tacacs',
				   'Serial Number'
	);
	$ws->write_row( $rows, 0, \@header, $fmtGray );
	$rows++;

	# CIS values - cistype{MLXS}{router} = value
	my %cistype = (
			   'M', { '2951', 'C2951', '3945E', 'C3945E', 'ASR',   'ASR1001-X', '4321', 'C4321', '4351', 'C4351', '4451', '4451', '4461', 'C4651', '4331', 'C4331' },
			   'L', { '2951', 'C2951', '3945E', 'C3945E', 'ASR',   'ASR1001-X', '4321', 'C4321', '4351', 'C4351', '4451', 'C4451', '4461', 'C4461' },
			   'X', { '2951', 'C2951', '3945E', 'C3945E', 'ASR',   'ASR1001-X', '4321', 'C4321', '4351', 'C4351', '4451', 'C4451', '4461', 'C4461' },
			   'S', { '2951', 'C2951', '3945',  'C3945',  '3945E', 'C3945E',    '4321', 'C4321', '4351', 'C4351', '4451', 'C4451', '4331', 'C4331' },
	);
	my ( $cistype, $cistype2 );
	if ( $sitetype =~ /([MLSX])/ ) {
		$cistype = $cistype{$1}{ $anchor{'router_type'} };
		$cistype2 = $cistype;
	} elsif ( $sitetype eq 'P' ) {
		$cistype = 'C4331';    # changed from c4451 - 07/16/2020
	} elsif ( $sitetype eq 'Q' ) {
		$cistype = 'C2951';
	}
	my ( $mlstype, $mlsupl, $mlsupl1, $mlsupl2, $mlsinter );

	my %mlstype = ( 'M', 'C9300-48T-A', 'L', 'C9407R', 'X', 'C9407R' );
	$mlstype = $mlstype{$sitetype} if ( defined $mlstype{$sitetype} );
	if ( $sitetype eq 'M' ) {
		$mlsupl   = '-g1/0/';
		$mlsupl1  = '-g1/0/';
		$mlsupl2  = '-g2/0/';
		$mlsinter = '-fa1/0/';
	} elsif ( $sitetype eq 'L' or $sitetype eq 'X' ) {
		$mlsupl   = '-g5/';
		$mlsinter = '-g5/';
	}
	$mlsupl =~ tr/\//-/;

	# Decide whether Voice VLANs should be entered as DHCP scopes or not
	my $vdhcp = 'No';
	$vdhcp = 'Yes*' if ( $anchor{'ipt'} eq 'Y' );

	# Start writing output
	my $wantype = $anchor{'pri_circuit_type'};
	my @row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', '', 'No', 'Yes', 'Yes', 'Yes', 'Yes', '' );
	if (    $sitetype eq 'X'
		 or $sitetype eq 'L'
		 or $sitetype eq 'M'
		 or $sitetype eq 'S' )
	{
		$row[5] = $anchor{'tad1_subnet'};    # $anchor{'loop_subnet'} . '.160/30';
		$row[6] = $anchor{'tad1_ip'};        # $anchor{'loop_subnet'} . '.161';
		$row[7] = $anchor{'tad1_name'};
		$row[8] = 'SLC8000';
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	if ( $sitetype eq 'P' ) {
		@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', '', 'No', 'Yes', 'No', 'No', 'No', '' );
		$row[5] = $anchor{'loop_subnet'} . '.160/30';
		$row[6] = $anchor{'loop_subnet'} . '.161';
		$row[7] = 'tad' . $siteid . $anchor{'mdf_flr'} . 'a01';
		$row[8] = 'SLC8000';
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
		@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', '', 'No', 'Yes', 'No', 'No', 'No', '' );
	}
	if ( $anchor{'cis1_name'} ne '' ) {
		$row[5] = $anchor{'loop_subnet'} . '.1/32';
		$row[6] = $anchor{'loop_subnet'} . '.1';
		$row[7] = $anchor{'cis1_name'};
		$row[8] = $cistype;
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	if ( $anchor{'cis2_name'} ne '' ) {
		$row[5] = $anchor{'loop_subnet'} . '.2/32';
		$row[6] = $anchor{'loop_subnet'} . '.2';
		$row[7] = $anchor{'cis2_name'};
		$row[8] = $cistype2;
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	if ( $anchor{'mls1_name'} ne '' and $sitetype =~ /^[XLM]$/ ) {
		$row[5] = $anchor{'loop_subnet'} . '.33/32';
		$row[6] = $anchor{'loop_subnet'} . '.33';
		$row[7] = $anchor{'mls1_name'};
		$row[8] = $mlstype;
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	if ( $sitetype !~ /^[MS]$/ and $anchor{'mls2_name'} ne '' ) {
		$row[5] = $anchor{'loop_subnet'} . '.34/32';
		$row[6] = $anchor{'loop_subnet'} . '.34';
		$row[7] = $anchor{'mls2_name'};
		$row[8] = $mlstype;
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}

	# VGs

	# Making the VG output as site-type agnostic as possible. I think that will work best for future changes.
	if ( $anchor{'vgcount'} == 1 ) {
		@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', 'VG204XM', 'No', 'Yes', 'Yes', 'Yes', 'Yes', '' );

			$row[5] = $anchor{'vgc1_address'} . '/24';
			$row[6] = $anchor{'vgc1_address'};
			$row[7] = $anchor{'vgc1_name'};
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

	} elsif ( $anchor{'vgcount'} == 2 ) {
		@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', 'VG204XM', 'No', 'Yes', 'Yes', 'Yes', 'Yes', '' );

		$row[5] = $anchor{'vgc1_address'} . '/24';
		$row[6] = $anchor{'vgc1_address'};
		$row[7] = $anchor{'vgc1_name'};
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

		$row[5] = $anchor{'vgc2_address'} . '/24';
		$row[6] = $anchor{'vgc2_address'};
		$row[7] = $anchor{'vgc2_name'};
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
	}

	if ( $anchor{'pri_wan_ip_cer'} ne '' ) {
		@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', '', 'na', 'Yes', 'No', 'No', 'No', '' );
		( my $octet1, my $octet2, my $octet3, my $octet4 ) =
		  split( /\./, $anchor{'pri_wan_ip_cer'} );
		$octet4--;
		my $wan_subnet = join( '.', $octet1, $octet2, $octet3, $octet4 );

	if ( $anchor{'r1_provider'} eq 'ATT' and $anchor{'transport'} == 3 and $anchor{'SDWAN'} eq 'Y' ){
			$row[5] = $wan_subnet . '/30';
			$row[6] = $anchor{'pri_wan_ip_cer'};
			$row[7] = $anchor{'cis1_name'} . '-' . $anchor{'att_upl_int_dns'};
			$row[8] = 'e-PVC';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $wan_subnet . '/30';
			$row[6] = $anchor{'pri_wan_ip_per'};
			$row[7] = 'per-' . $anchor{'cis1_name'};
			$row[8] = 'e-PVC';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		# Make router 2 show up in the IP Summ Chart if it has the AT&T circuit
	}elsif( $anchor{'r2_provider'} eq 'ATT' ){

		#Create a locally scoped value for $wan_subnet to avoid it picking up the WAN IP subnet of R1
		( my $octet1, my $octet2, my $octet3, my $octet4 ) = split( /\./, $anchor{'sec_wan_ip_cer'} );
		$octet4--;
		my $wan_subnet = join( '.', $octet1, $octet2, $octet3, $octet4 );

			$row[5] = $wan_subnet . '/30';
			$row[6] = $anchor{'sec_wan_ip_cer'};
			$row[7] = $anchor{'cis2_name'} . '-' . $anchor{'att_upl_int_dns'};
			$row[8] = 'e-PVC';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $wan_subnet . '/30';
			$row[6] = $anchor{'sec_wan_ip_per'};
			$row[7] = 'per-' . $anchor{'cis2_name'};
			$row[8] = 'e-PVC';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}
	}
	if ( $sitetype eq 'M' ) {
		if ( $anchor{'SDWAN'} ne 'Y' ) {
			@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', '', 'No', 'Yes', 'Yes', 'Yes', 'Yes', '' );
			$row[5] = $anchor{'loop_subnet'} . '.128/30';
			$row[6] = $anchor{'loop_subnet'} . '.129';
			$row[7] = $anchor{'cis1_name'} . '-g0-0-1';
			$row[8] = 'cis1 to mls1-1';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.128/30';
			$row[6] = $anchor{'loop_subnet'} . '.130';
			$row[7] = $anchor{'mls1_name'} . $mlsupl1 . '1';
			$row[8] = 'mls1-1 to cis1';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.132/30';
			$row[6] = $anchor{'loop_subnet'} . '.133';
			$row[7] = $anchor{'cis1_name'} . '-g0-0-2';
			$row[8] = 'cis1 to mls1-2';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.132/30';
			$row[6] = $anchor{'loop_subnet'} . '.134';
			$row[7] = $anchor{'mls1_name'} . $mlsupl2 . '2';
			$row[8] = 'mls1-2 to cis1';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.136/30';
			$row[6] = $anchor{'loop_subnet'} . '.137';
			$row[7] = $anchor{'cis2_name'} . '-g0-0-1';
			$row[8] = 'cis2 to mls1-2';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.136/30';
			$row[6] = $anchor{'loop_subnet'} . '.138';
			$row[7] = $anchor{'mls1_name'} . $mlsupl2 . '1';
			$row[8] = 'mls1-2 to cis2';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.140/30';
			$row[6] = $anchor{'loop_subnet'} . '.141';
			$row[7] = $anchor{'cis2_name'} . '-g0-0-2';
			$row[8] = 'cis2 to mls1-1';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.140/30';
			$row[6] = $anchor{'loop_subnet'} . '.142';
			$row[7] = $anchor{'mls1_name'} . $mlsupl1 . '2';
			$row[8] = 'mls1-1 to cis2';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}elsif ( $anchor{'SDWAN'} eq 'Y' ) {
			@row = (
					 'New', '', '', '',
					 'mdf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.65',
					 $anchor{'cis1_name'} . '-g0-0-1-100',
					 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '',
					 'mdf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.66',
					 $anchor{'cis2_name'} . '-g0-0-1-100',
					 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '',
					 'Vlan100/idf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.72',
					 $anchor{'mls1_name'} . '-vlan-100',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '',
					 'mdf',
					 $anchor{'loop_subnet'} . '.48/30',
					 $anchor{'loop_subnet'} . '.49',
					 $anchor{'cis1_name'} . '-g0-1-0-40',
					 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '',
					 'mdf',
					 $anchor{'loop_subnet'} . '.48/30',
					 $anchor{'loop_subnet'} . '.50',
					 $anchor{'cis2_name'} . '-g0-0-2-40',
					 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}

		if ( $anchor{'fw'} eq 'Y' ) {
			@row = ( 'New', 'na', '', 'na', 'Vlan99/mdf', '', '', '', '', 'na', 'Yes', 'No', 'No', 'No', '' );
			$row[4] = 'Vlan99/mdf';
			$row[5] = $anchor{'loop_subnet'} . '.80/28';
			$row[6] = $anchor{'loop_subnet'} . '.81';
			$row[7] = $anchor{'mls1_name'} . '-gwy-vlan-99';
			$row[8] = 'Firewall Inside Vlan GWY';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}
		$row[4] = 'Vlan101/mdf';
		$row[5] = $anchor{'svr_subnet_1'} . '.0/24';
		$row[6] = $anchor{'svr_subnet_1'} . '.1';
		$row[7] = $anchor{'mls1_name'} . '-gwy-vlan-101';
		$row[8] = '1st Server Vlan GWY';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

		# Stacks
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack        = $ct + 1;
			my $stknum       = sprintf( "%02d", $ct + 1 );
			my $stacksubnet  = 'idf' . $stack . '_ds_1';
			my $stackvsubnet = 'idf' . $stack . '_vs_1';
			my $stackesubnet = 'idf' . $stack . '_es_1';
			my $stackuplink  = 'idf' . $stack . '_mls_up';
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};
			$stackuplink  = $anchor{$stackuplink};
			$stackuplink =~ tr/\//-/;
			$stknum = sprintf( "%02d", $stknum );
			@row = (
					 'New', '', '', '', 'Vlan2' . $stknum . '/idf',
					 '',
					 $stacksubnet . '.1',
					 $anchor{'mls1_name'} . '-g' . $stackuplink . '-gwy',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[4] = 'Vlan3' . $stknum . '/idf';
			$row[6] = $stackvsubnet . '.1';
			$row[7] = $anchor{'mls1_name'} . '-g' . $stackuplink . '-gwy-ipt';
			$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
			$row[4] = 'Vlan4' . $stknum . '/idf';
			$row[6] = $stackesubnet . '.1';
			$row[7] = $anchor{'mls1_name'} . '-g' . $stackuplink . '-gwy-eac';
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );
		}
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack = $ct + 1;
			( my $stackname, my $stknum ) = split( /,/, $StackList[$ct] );
			my $stacksubnet  = 'idf' . $stack . '_ds_1';
			my $stackvsubnet = 'idf' . $stack . '_vs_1';
			my $stackesubnet = 'idf' . $stack . '_es_1';
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};
			@row = (
					 'New', 'na', $stknum,
					 $stknum * 48,
					 'IDF Stack' . $stack . '- Data',
					 $stacksubnet . '.0/24',
					 $stacksubnet . '.4',
					 $stackname, 'C9300-48P-E', 'Yes', 'Yes', 'Yes', 'Yes', 'Yes', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtGreen );
			@row = (
					 'New', 'na', $stknum,
					 $stknum * 48,
					 'IDF Stack' . $stack . '- Voice',
					 $stackvsubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, 'C9300-48P-E', $vdhcp, 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
			@row = (
					 'New', 'na', $stknum,
					 $stknum * 48,
					 'IDF Stack' . $stack . '- EAC',
					 $stackesubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, 'C9300-48P-E', 'Yes', 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );
		}
		if ( $anchor{'wlan'} eq 'Y' ) {
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . '.0/24',
					 $anchor{'wlan_subnet_i'} . '.1',
					 $anchor{'mls1_name'} . '-vlan-110-gwy',
					 'WLAN User VLAN-HSRP',
					 'Yes', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . '.0/24',
					 $anchor{'wlan_subnet_i'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-110',
					 'WLC5500', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} ne '9800' );
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . '.0/24',
					 $anchor{'wlan_subnet_i'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-110',
					 'WLC9800', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} eq '9800' );
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.1',
					 $anchor{'mls1_name'} . '-vlan-119-gwy',
					 'WLAN EAC VLAN-HSRP',
					 'Yes*  **', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-119',
					 'WLC5500', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} ne '9800' );
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-119',
					 'WLC9800', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} eq '9800' );
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan101/mdf',
					 $anchor{'svr_subnet_1'} . '.0/24',
					 $anchor{'svr_subnet_1'} . '.9',
					 $anchor{'wlc1_name'}, 'WLC5500', 'No', 'Yes', 'No', 'Yes', 'No', ''
			) if ( $anchor{'wlc_model'} ne '9800' );
			@row = (
					 'New', '', '', '', 'Vlan101/mdf',
					 $anchor{'svr_subnet_1'} . '.0/24',
					 $anchor{'svr_subnet_1'} . '.9',
					 $anchor{'wlc1_name'}, 'WLC9800', 'No', 'Yes', 'No', 'Yes', 'No', ''
			) if ( $anchor{'wlc_model'} eq '9800' );
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
		}
	} elsif ( $sitetype eq 'L' or $sitetype eq 'X' ) {
		@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', '', 'No', 'Yes', 'Yes', 'Yes', 'Yes', '' );
		$row[5] = $anchor{'loop_subnet'} . '.128/30';
		$row[6] = $anchor{'loop_subnet'} . '.129';
		$row[7] = $anchor{'cis1_name'} . '-g0-0';
		$row[8] = 'cis1 to mls1';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.128/30';
		$row[6] = $anchor{'loop_subnet'} . '.130';
		$row[7] = $anchor{'mls1_name'} . $mlsupl . '1';
		$row[8] = 'mls1 to cis1';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.132/30';
		$row[6] = $anchor{'loop_subnet'} . '.133';
		$row[7] = $anchor{'cis1_name'} . '-g0-2';
		$row[8] = 'cis1 to mls2';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.132/30';
		$row[6] = $anchor{'loop_subnet'} . '.134';
		$row[7] = $anchor{'mls2_name'} . $mlsupl . '2';
		$row[8] = 'mls2 to cis1';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.136/30';
		$row[6] = $anchor{'loop_subnet'} . '.137';
		$row[7] = $anchor{'cis2_name'} . '-g0-0';
		$row[8] = 'cis2 to mls2';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.136/30';
		$row[6] = $anchor{'loop_subnet'} . '.138';
		$row[7] = $anchor{'mls2_name'} . $mlsupl . '1';
		$row[8] = 'mls2 to cis2';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.140/30';
		$row[6] = $anchor{'loop_subnet'} . '.141';
		$row[7] = $anchor{'cis2_name'} . '-g0-2';
		$row[8] = 'cis2 to mls1';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.140/30';
		$row[6] = $anchor{'loop_subnet'} . '.142';
		$row[7] = $anchor{'mls1_name'} . $mlsupl . '2';
		$row[8] = 'mls1 to cis2';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

		if ( $anchor{'fw'} eq 'Y' ) {
			@row = ( 'New', 'na', '', 'na', 'Vlan99/mdf', '', '', '', '', 'na', 'Yes', 'No', 'No', 'No', '' );
			$row[5] = $anchor{'loop_subnet'} . '.80/28';
			$row[6] = $anchor{'loop_subnet'} . '.81';
			$row[7] = $anchor{'mls1_name'} . '-hsrp-vlan-99';
			$row[8] = 'Firewall Inside Vlan HSRP';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.80/28';
			$row[6] = $anchor{'loop_subnet'} . '.82';
			$row[7] = $anchor{'mls1_name'} . '-vlan-99';
			$row[8] = 'Firewall Inside Vlan';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[5] = $anchor{'loop_subnet'} . '.80/28';
			$row[6] = $anchor{'loop_subnet'} . '.83';
			$row[7] = $anchor{'mls2_name'} . '-vlan-99';
			$row[8] = 'Firewall Inside Vlan';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}

		#Add TLOC Intefaces for Large SDWAN sites
		if ( $anchor{'SDWAN'} eq 'Y'){
			@row = ( 'New', '', '', '', 'mdf', '', '', '', '', 'na', 'Yes', 'No', 'No', 'No', '' );
			$row[4] = 'mdf';
			$row[5] = $anchor{'loop_subnet'} . '.48/30';
			$row[6] = $anchor{'loop_subnet'} . '.49';
			$row[7] = $anchor{'cis1_name'} . '-g0-1-0-40';
			$row[8] = $cistype;
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

			@row = ( 'New', '', '', '', 'mdf', '', '', '', '', 'na', 'Yes', 'No', 'No', 'No', '' );
			$row[4] = 'mdf';
			$row[5] = $anchor{'loop_subnet'} . '.48/30';
			$row[6] = $anchor{'loop_subnet'} . '.50';
			$row[7] = $anchor{'cis2_name'} . '-g0-0-02-40';
			$row[8] = $cistype;
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}

		@row = ( 'New', 'na', '', 'na', 'mdf', '', '', '', '', 'na', 'Yes', 'No', 'No', 'No', '' );
		$row[4] = 'Vlan100/mdf';
		$row[5] = $anchor{'loop_subnet'} . '.144/30';
		$row[6] = $anchor{'loop_subnet'} . '.145';
		$row[7] = $anchor{'mls1_name'} . '-vlan-100';
		$row[8] = 'mls1 to mls2';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'loop_subnet'} . '.144/30';
		$row[6] = $anchor{'loop_subnet'} . '.146';
		$row[7] = $anchor{'mls2_name'} . '-vlan-100';
		$row[8] = 'mls2 to mls1';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[4] = 'Vlan101/mdf';
		$row[5] = $anchor{'svr_subnet_1'} . '.0/24';
		$row[6] = $anchor{'svr_subnet_1'} . '.1';
		$row[7] = $anchor{'mls1_name'} . '-hsrp-vlan-101';
		$row[8] = '1st Server Vlan HSRP';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'svr_subnet_1'} . '.0/24';
		$row[6] = $anchor{'svr_subnet_1'} . '.2';
		$row[7] = $anchor{'mls1_name'} . '-vlan-101';
		$row[8] = '1st Server Vlan';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		$row[5] = $anchor{'svr_subnet_1'} . '.0/24';
		$row[6] = $anchor{'svr_subnet_1'} . '.3';
		$row[7] = $anchor{'mls2_name'} . '-vlan-101';
		$row[8] = '1st Server Vlan';
		$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

		# VLAN
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack        = $ct + 1;
			my $stknum       = sprintf( "%02d", $ct + 1 );
			my $stacksubnet  = 'idf' . $stack . '_ds_1';
			my $stackvsubnet = 'idf' . $stack . '_vs_1';
			my $stackesubnet = 'idf' . $stack . '_es_1';
			my $stackuplink  = 'idf' . $stack . '_mls_up';
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};
			$stackuplink  = $anchor{$stackuplink};
			$stackuplink =~ tr/\//-/;

			# Default to odd numbered stacks
			my @mlsname = (
							'mls1_name', 'mls1_name', 'mls2_name', 'mls2_name', 'mls2_name', 'mls1_name',
							'mls1_name', 'mls1_name', 'mls2_name'
			);
			if ( $stknum % 2 == 0 ) {    # switch order for even numbered stacks
				@mlsname = (
							 'mls2_name', 'mls2_name', 'mls1_name', 'mls1_name', 'mls1_name', 'mls2_name',
							 'mls2_name', 'mls2_name', 'mls1_name'
				);
			}
			my $mlsname = shift @mlsname;
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan2' . $stknum . '/idf',
					 '',
					 $stacksubnet . '.1',
					 $mlsname . '-g' . $stackuplink . '-gwy',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtBlue );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan2' . $stknum . '/idf',
					 '',
					 $stacksubnet . '.2',
					 $mlsname . '-g' . $stackuplink,
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtBlue );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan2' . $stknum . '/idf',
					 '',
					 $stacksubnet . '.3',
					 $mlsname . '-g' . $stackuplink,
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtBlue );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan3' . $stknum . '/idf',
					 '',
					 $stackvsubnet . '.1',
					 $mlsname . '-g' . $stackuplink . '-gwy-ipt',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtOrange );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan3' . $stknum . '/idf',
					 '',
					 $stackvsubnet . '.2',
					 $mlsname . '-g' . $stackuplink . '-ipt',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtOrange );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan3' . $stknum . '/idf',
					 '',
					 $stackvsubnet . '.3',
					 $mlsname . '-g' . $stackuplink . '-ipt',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtOrange );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan4' . $stknum . '/idf',
					 '',
					 $stackesubnet . '.1',
					 $mlsname . '-g' . $stackuplink . '-gwy-eac',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtTeal );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan4' . $stknum . '/idf',
					 '',
					 $stackesubnet . '.2',
					 $mlsname . '-g' . $stackuplink . '-eac',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows    = writeRow( $ws, $rows, \@row, $fmtTeal );
			$mlsname = shift(@mlsname);
			$mlsname = $anchor{$mlsname};
			@row = (
					 'New', '', '', '', 'Vlan4' . $stknum . '/idf',
					 '',
					 $stackesubnet . '.3',
					 $mlsname . '-g' . $stackuplink . '-eac',
					 $mlstype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );
		}

		# Stacks
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack = $ct + 1;
			( my $stackname, my $stknum ) = split( /,/, $StackList[$ct] );
			my $stacksubnet  = 'idf' . $stack . '_ds_1';
			my $stackvsubnet = 'idf' . $stack . '_vs_1';
			my $stackesubnet = 'idf' . $stack . '_es_1';
			my $stackuplink  = 'idf' . $stack . '_mls_up';
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};
			$stackuplink  = $anchor{$stackuplink};
			$stackuplink =~ tr/\//-/;
			@row = (
					 'New', 'na', $stknum,
					 $stknum * 48,
					 'IDF Stack' . $stack . '- Data',
					 $stacksubnet . '.0/24',
					 $stacksubnet . '.4',
					 $stackname, 'C9300-48P-E', 'Yes', 'Yes', 'Yes', 'Yes', 'Yes', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtGreen );
			@row = (
					 'New', 'na', $stknum,
					 $stknum * 48,
					 'IDF Stack' . $stack . '- Voice',
					 $stackvsubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, 'C9300-48P-E', $vdhcp, 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
			@row = (
					 'New', 'na', $stknum,
					 $stknum * 48,
					 'IDF Stack' . $stack . '- EAC',
					 $stackesubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, 'C9300-48P-E', 'Yes', 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );
		}
		if ( $anchor{'wlan'} eq 'Y' ) {

			# The mask changes between site type X and L for the first 5 rows - nothing else changes
			my $mask = '.0/22';    # Site type X
			$mask = '.0/24' if ( $sitetype eq 'L' );    # Site type L
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . $mask,
					 $anchor{'wlan_subnet_i'} . '.1',
					 $anchor{'mls2_name'} . '-vlan-110-gwy',
					 'WLAN User VLAN-HSRP',
					 'Yes', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . $mask,
					 $anchor{'wlan_subnet_i'} . '.2',
					 $anchor{'mls2_name'} . '-vlan-110',
					 'WLAN User VLAN',
					 'Yes', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . $mask,
					 $anchor{'wlan_subnet_i'} . '.3',
					 $anchor{'mls1_name'} . '-vlan-110',
					 'WLAN User VLAN',
					 'Yes', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . $mask,
					 $anchor{'wlan_subnet_i'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-110',
					 'WLC5500', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} ne '9800' );
			@row = (
					 'New', '', '', '', 'Vlan110/mdf',
					 $anchor{'wlan_subnet_i'} . $mask,
					 $anchor{'wlan_subnet_i'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-110',
					 'WLC9800', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} eq '9800' );
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );

			# Masks for the rest are the same with L and X sites, .0/24
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.1',
					 $anchor{'mls1_name'} . '-vlan-119-gwy',
					 'WLAN EAC VLAN-HSRP',
					 'Yes*  **', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.2',
					 $anchor{'mls1_name'} . '-vlan-119',
					 'WLAN EAC VLAN',
					 'Yes*  **', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.3',
					 $anchor{'mls2_name'} . '-vlan-119',
					 'WLAN EAC VLAN',
					 'Yes*  **', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-119',
					 'WLC5500', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} ne '9800' );
			@row = (
					 'New', '', '', '', 'Vlan119/mdf',
					 $anchor{'wlan_subnet_e'} . '.0/24',
					 $anchor{'wlan_subnet_e'} . '.4',
					 $anchor{'wlc1_name'} . '-vlan-119',
					 'WLC9800', 'No', 'Yes', 'No', 'No', 'No', ''
			) if ( $anchor{'wlc_model'} eq '9800' );
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
			@row = (
					 'New', '', '', '', 'Vlan101/mdf',
					 $anchor{'svr_subnet_1'} . '.0/24',
					 $anchor{'svr_subnet_1'} . '.9',
					 $anchor{'wlc1_name'}, 'WLC5500', 'No', 'Yes', 'No', 'Yes', 'No', ''
			) if ( $anchor{'wlc_model'} ne '9800' );
			@row = (
					 'New', '', '', '', 'Vlan101/mdf',
					 $anchor{'svr_subnet_1'} . '.0/24',
					 $anchor{'svr_subnet_1'} . '.9',
					 $anchor{'wlc1_name'}, 'WLC9800', 'No', 'Yes', 'No', 'Yes', 'No', ''
			) if ( $anchor{'wlc_model'} eq '9800' );
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
		}
	} elsif ( $sitetype eq 'S' ) {

		my $stackmodel;
		if ( $anchor{'stack_vendor'} eq 'aruba' ){
			$stackmodel = '6300';
		}else{
			$stackmodel = 'C9300-48P-E';
		}

		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			( my $stackname, my $stknum ) = split( /,/, $StackList[$ct] );
			my $stack        = $ct + 1;
			$stknum       = sprintf( "%02d", $ct + 1 );
			my $stacksubnet  = 'data_subnet_' . $stack;
			my $stackvsubnet = 'voice_subnet_' . $stack;
			my $stackesubnet = 'eac_subnet_' . $stack;
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};
			@row = (
					 'New', '', '', '', 'mdf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.65',
					 $anchor{'cis1_name'} . '-g0-0-1-100',
					 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '', 'mdf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.66',
					 $anchor{'cis2_name'} . '-g0-0-1-100',
					 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '', 'Vlan100/idf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.72',
					 $stackname . '-vlan-100',
					 $stackmodel,
					 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );


			# Allows TLOCs to show up in IP Summary Chart if site is SDWAN
			if ( $anchor{'SDWAN'} eq 'Y' and $anchor{'transport'} == 2 ){
				@row = (
						 'New', '', '', '', 'mdf',
						 $anchor{'loop_subnet'} . '.48/30',
						 $anchor{'loop_subnet'} . '.49',
						 $anchor{'cis1_name'} . '-g0-0-2-40',
						 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
				);
				$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
				@row = (
						 'New', '', '', '', 'mdf',
						 $anchor{'loop_subnet'} . '.48/30',
						 $anchor{'loop_subnet'} . '.50',
						 $anchor{'cis2_name'} . '-g0-0-2-40',
						 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
				);
				$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
				@row = (
						 'New', '', '', '', 'mdf',
						 $anchor{'loop_subnet'} . '.52/30',
						 $anchor{'loop_subnet'} . '.53',
						 $anchor{'cis1_name'} . '-g0-0-2-20',
						 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
				);
				$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
				@row = (
						 'New', '', '', '', 'mdf',
						 $anchor{'loop_subnet'} . '.52/30',
						 $anchor{'loop_subnet'} . '.54',
						 $anchor{'cis2_name'} . '-g0-0-2-20',
						 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
				);
				$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			}elsif( $anchor{'SDWAN'} eq 'Y' and $anchor{'transport'} == 3 ){
				@row = (
						 'New', '', '', '', 'mdf',
						 $anchor{'loop_subnet'} . '.48/30',
						 $anchor{'loop_subnet'} . '.49',
						 $anchor{'cis1_name'} . '-g0-0-2-40',
						 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
				);
				$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
				@row = (
						 'New', '', '', '', 'mdf',
						 $anchor{'loop_subnet'} . '.48/30',
						 $anchor{'loop_subnet'} . '.50',
						 $anchor{'cis2_name'} . '-g0-0-2-40',
						 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
				);
				$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			}

			@row = (
					 'New', '', '', '', 'Vlan201' . '/idf',
					 '',
					 $stacksubnet . '.1',
					 $stackname . '-vlan-201-gwy',
					 $stackmodel,
					  'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '', 'Vlan301' . '/idf',
					 '',
					 $stackvsubnet . '.1',
					 $stackname . '-vlan-301-gwy-ipt',
					 $stackmodel,
					 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
			@row = (
					 'New', '', '', '', 'Vlan901' . '/idf',
					 $anchor{'loop_subnet'} . '.160/30',,
					 $anchor{'loop_subnet'} . '.162',
					 $stackname . '-vlan-901',
					 $stackmodel,
					 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
			@row = (
					 'New', '', '', '', 'Vlan401' . '/idf',
					 '',
					 $stackesubnet . '.1',
					 $stackname . '-vlan-401-gwy-eac',
					 $stackmodel,
					 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );

		}
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack = $ct + 1;
			( my $stackname, my $stknum ) = split( /,/, $StackList[$ct] );
			my $stacksubnet  = 'data_subnet_' . $stack;
			my $stackvsubnet = 'voice_subnet_' . $stack;
			my $stackesubnet = 'eac_subnet_' . $stack;
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};

			# The math is different for a 24-port switch
			my $stackamt = $stknum;
			$stackamt = 0.5 if ( $stknum eq '24-port' );
			@row = (
					 'New', 'na', $stknum,
					 $stackamt * 48,
					 'IDF Stack' . $stack . '- Data',
					 $stacksubnet . '.0/24',
					 $stacksubnet . '.4',
					 $stackname, $stackmodel, 'Yes', 'Yes', 'Yes', 'Yes', 'Yes', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtGreen );
			@row = (
					 'New', 'na', $stknum,
					 $stackamt * 48,
					 'IDF Stack' . $stack . '- Voice',
					 $stackvsubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, $stackmodel, $vdhcp, 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
			@row = (
					 'New', 'na', $stknum,
					 $stackamt * 48,
					 'IDF Stack' . $stack . '- EAC',
					 $stackesubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, $stackmodel, 'Yes', 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );

			#wlc-virtual controller for small sites
			if ( $anchor{'stack_vendor'} eq 'aruba' ) {

				@row = (
					 'New', '', '', '', 'IDF Stack' . $stack . '- Data',
					 $stacksubnet . '.0/24',
					 $stacksubnet . '.9',
					 'wlc' . substr($stackname, 3) . '-vc',
					 'Virtual-Controller',
					 'na', 'Yes', 'No', 'No', 'No', ''
					);
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );

			}
		}
	} elsif ( $sitetype eq 'P' or $sitetype eq 'Q' ) {
			for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			( my $stackname, my $stknum ) = split( /,/, $StackList[$ct] );
			my $stack        = $ct + 1;
			$stknum       = sprintf( "%02d", $ct + 1 );
			my $stacksubnet  = 'data_subnet_' . $stack;
			my $stackvsubnet = 'voice_subnet_' . $stack;
			my $stackesubnet = 'eac_subnet_' . $stack;
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};
			@row = (
					 'New', '', '', '', 'mdf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.65',
					 $anchor{'cis1_name'} . '-g0-0-1-100',
					 $cistype, 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '', 'Vlan100/idf',
					 $anchor{'loop_subnet'} . '.64/28',
					 $anchor{'loop_subnet'} . '.72',
					 $stackname . '-vlan-100',
					 'C9300-48P-E', 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

			@row = (
					 'New', '', '', '', 'Vlan201' . '/idf',
					 '',
					 $stacksubnet . '.1',
					 $stackname . '-vlan-201-gwy',
					 'C9300-48P-E', 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
			@row = (
					 'New', '', '', '', 'Vlan301' . '/idf',
					 '',
					 $stackvsubnet . '.1',
					 $stackname . '-vlan-301-gwy-ipt',
					 'C9300-48P-E', 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
			@row = (
					 'New', '', '', '', 'Vlan901' . '/idf',
					 $anchor{'loop_subnet'} . '.160/30',,
					 $anchor{'loop_subnet'} . '.162',
					 $stackname . '-vlan-901',
					 'C9300-48P-E', 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
			@row = (
					 'New', '', '', '', 'Vlan401' . '/idf',
					 '',
					 $stackesubnet . '.1',
					 $stackname . '-vlan-401-gwy-eac',
					 'C9300-48P-E', 'na', 'Yes', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );
		}
			for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack = $ct + 1;
			( my $stackname, my $stknum ) = split( /,/, $StackList[$ct] );
			my $stacksubnet  = 'data_subnet_' . $stack;
			my $stackvsubnet = 'voice_subnet_' . $stack;
			my $stackesubnet = 'eac_subnet_' . $stack;
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};

			# The math is different for a 24-port switch
			my $stackamt = $stknum;
			$stackamt = 0.5 if ( $stknum eq '24-port' );
			@row = (
					 'New', 'na', $stknum,
					 $stackamt * 48,
					 'IDF Stack' . $stack . '- Data',
					 $stacksubnet . '.0/24',
					 $stacksubnet . '.4',
					 $stackname, 'C9300-48P-E', 'Yes', 'Yes', 'Yes', 'Yes', 'Yes', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtGreen );
			@row = (
					 'New', 'na', $stknum,
					 $stackamt * 48,
					 'IDF Stack' . $stack . '- Voice',
					 $stackvsubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, 'C9300-48P-E', $vdhcp, 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
			@row = (
					 'New', 'na', $stknum,
					 $stackamt * 48,
					 'IDF Stack' . $stack . '- EAC',
					 $stackesubnet . '.0/24',
					 '(use data vlan)',
					 $stackname, 'C9300-48P-E', 'Yes', 'No', 'No', 'No', 'No', ''
			);
			$rows = writeRow( $ws, $rows, \@row, $fmtTeal );
		}
	} else {
		prtout("Site Model type '$sitetype' is not valid");
		xit(1);
	}
	$rows++;    # leave a blank line before the notes and legend
	@row = ( 'Key', '', '', '', 'if applicable---->', '* - Requires option 150 TFTP Servers' );
	$rows = writeRow( $ws, $rows, \@row );

	# Formats are different for the cells in this row, so write them directly instead of passing to the sub
	$ws->write( $rows, 0, 'Loopback/Mgmt', $fmtPurple );
	$ws->write( $rows, 5, '* - Please add option 150 for the voice DHCP scope per IPT Engineering Team Standards' );
	$rows++;
	$ws->write( $rows, 0, 'Router/MLS', $fmtBlue );
	$ws->write( $rows, 5, '** - Wireless LAN DHCP scopes should be configured with a 2-hour lease time' );
	$rows++;

	# Back to normal with formats now
	@row  = ('Stacks');
	$rows = writeRow( $ws, $rows, \@row, $fmtGreen );
	@row  = ('IPT');
	$rows = writeRow( $ws, $rows, \@row, $fmtOrange );
	@row  = ('EAC');
	$rows = writeRow( $ws, $rows, \@row, $fmtTeal );
	@row  = ('Wireless');
	$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
	return $outputFile;
}

sub writeCISummary {
	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid, my $sitetype, my $stacklimit ) = @_;
	my $outputFile = $siteid . '-CI-Devices.xls';

	my $workbook = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile")
	  or die "create XLS file '$SVR_ROOTDIR/$OutputDir/$outputFile' failed: $!";

	prtout("Writing CI Summary XLS");
	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75);
	$ws->set_column( 'A:A', 14.5 );
	$ws->set_column( 'B:B', 30 );
	$ws->set_column( 'C:C', 17 );
	$ws->set_column( 'D:D', 18 );
	$ws->set_column( 'E:E', 25 );
	$ws->set_column( 'F:F', 16.5 );
	$ws->set_column( 'G:G', 16.5 );
	$ws->set_column( 'H:H', 18 );
	$ws->set_column( 'I:I', 23 );
	$ws->set_column( 'J:J', 13 );
	$ws->set_column( 'K:K', 30 );
	$ws->set_column( 'L:L', 12 );
	$ws->set_column( 'M:M', 12 );
	my $fmtBlank = $workbook->add_format( size => 12, bold => 1 );
	my $fmtGray = $workbook->add_format(
										 size      => 9,
										 bold      => 1,
										 bg_color  => 22,
										 text_wrap => 1
	);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtYellow = $workbook->add_format( bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );

	# First three rows are the header
	my $rows   = 2;
	my $height = 12;
	$ws->merge_range( 'A1:O2', "$siteid - CI Summary Chart", $fmtBlank );
	my @header = (
				   'If Update to existing CI, provide HPSM CI ID',
				   'SYSTEM ID/FQDN (OVO ID)',
				   'FORMAL NAME',
				   'COMMON NAME',
				   'CI TYPE',
				   'ENVIRONMENT',
				   'SERIAL NUMBER',
				   'MODEL',
				   'BRAND',
				   'LOCATION',
				   'SUPPORTING WORKGROUP',
				   'REMARKS',
				   'IP ADDRESS',
				   'IP ADDRESS (additional'
	);
	$rows = writeRow( $ws, $rows, \@header, $fmtGray );

	# CIS values
	my %cistype = (
					'M', { '2951', 'C2951', '3945E', 'C3945E', 'ASR',   'ASR1001-X', '4321', 'C4321', '4351', 'C4351', '4331', 'C4331'},
					'L', { '2951', 'C2951', '3945E', 'C3945E', 'ASR',   'ASR1001-X', '4321', 'C4321', '4351', 'C4351' },
					'X', { '2951', 'C2951', '3945E', 'C3945E', 'ASR',   'ASR1001-X', '4321', 'C4321', '4351', 'C4351' },
					'S', { '2951', 'C2951', '3945',  'C3945',  '3945E', 'C3945E',    '4321', 'C4321', '4351', 'C4351', '4331','C4331'},
	);
	my ( $cistype, $cistype2 );
	if ( $sitetype =~ /([MLSX])/ ) {
		$cistype = $cistype{$1}{ $anchor{'router_type'} };
		$cistype2 = $cistype;
	} elsif ( $sitetype eq 'P' ) {
		$cistype = 'C4331';
	} elsif ( $sitetype eq 'Q' ) {
		$cistype = 'C2951';
	}
	my ( $mlstype, $mlsupl, $mlsupl1, $mlsupl2, $mlsinter );

	# Change 3750 to 3850 per Steve Groebe - 12/5/2016
	# change 3850 to 9300 per Sonja Mroczynski - 04/13/2020
	my %mlstype = ( 'M', 'C9300-48T-A', 'L', 'C9407R', 'X', 'C9407R' );
	$mlstype = $mlstype{$sitetype} if ( defined $mlstype{$sitetype} );
	if ( $sitetype eq 'M' ) {
		$mlsupl   = '-g1/0/';
		$mlsupl1  = '-g1/0/';
		$mlsupl2  = '-g2/0/';
		$mlsinter = '-fa1/0/';
	} elsif ( $sitetype eq 'L' or $sitetype eq 'X' ) {
		$mlsupl   = '-g4/';
		$mlsinter = '-g4/';
	}
	$mlsupl =~ tr/\//-/;

	# Decide whether Voice VLANs should be entered as DHCP scopes or not
	my $vdhcp = 'No';
	$vdhcp = 'Yes*' if ( $anchor{'ipt'} eq 'Y' );

	# Supporting workgroup is always the same
	my $workgroup = 'NHS_SLO_Remote'; # Per Sonja 04/30/2020

	# Write the output
	my @row = ( '', '', '', '', '', 'Production', '', '', 'Cisco', $siteid, $workgroup, '', '', '' );
	if ( $sitetype =~ /^(?:X|L|M|S)$/ ) {
		$row[2]  = $anchor{'tad1_name'};
		$row[1]  = $row[2] . '.uhc.com';
		$row[4]  = deviceType( $row[2] );
		$row[7]  = '';
		$row[12] = $anchor{'tad1_ip'};
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
		if ( $anchor{'vgcount'} == 1 ) {
			$row[2]  = $anchor{'vgc1_name'};
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'VG204';
			$row[12] = $anchor{'vgc1_address'};
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}
		elsif
		( $anchor{'vgcount'} == 2 ) {
			$row[2]  = $anchor{'vgc1_name'};
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'VG204';
			$row[12] = $anchor{'vgc1_address'};
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );

			$row[2]  = $anchor{'vgc2_name'};
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'VG204';
			$row[12] = $anchor{'vgc2_address'};
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}
	}
	if ( $anchor{'cis1_name'} ne '' ) {
		$row[2]  = $anchor{'cis1_name'};
		$row[1]  = $row[2] . '.uhc.com';
		$row[4]  = deviceType( $row[2] );
		$row[7]  = $anchor{'router_type'};
		$row[12] = $anchor{'loop_subnet'} . '.1';
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	if ( $anchor{'cis2_name'} ne '' ) {
		$row[2]  = $anchor{'cis2_name'};
		$row[1]  = $row[2] . '.uhc.com';
		$row[4]  = deviceType( $row[2] );
		$row[7]  = $anchor{'router_type'};
		$row[12] = $anchor{'loop_subnet'} . '.2';
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	if ( $anchor{'mls1_name'} ne '' and $sitetype =~ /^[XLM]$/ ) {
		$row[2]  = $anchor{'mls1_name'};
		$row[1]  = $row[2] . '.uhc.com';
		$row[4]  = deviceType( $row[2] );
		$row[7]  = $mlstype;
		$row[12] = $anchor{'loop_subnet'} . '.33';
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	# if ( $sitetype ne 'M' and $anchor{'mls2_name'} ne '' and $anchor{'mls2_name'} !~ 'stk') {
	if ( $anchor{'mls2_name'} ne '' and $sitetype =~ /^[XL]$/ ){
		print "Line 3635 : mls2_name has a value of : $anchor{'mls2_name'} \n";
		$row[2]  = $anchor{'mls2_name'};
		$row[1]  = $row[2] . '.uhc.com';
		$row[4]  = deviceType( $row[2] );
		$row[7]  = $mlstype;
		$row[12] = $anchor{'loop_subnet'} . '.34';
		$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	}
	if ( $sitetype eq 'M' ) {
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack        = $ct + 1;
			my $stackname    = ( split( /,/, $StackList[$ct] ) )[0];
			my $stacksubnet  = 'idf' . $stack . '_ds_1';
			my $stackvsubnet = 'idf' . $stack . '_vs_1';
			my $stackesubnet = 'idf' . $stack . '_es_1';
			$stacksubnet  = $anchor{$stacksubnet};
			$stackvsubnet = $anchor{$stackvsubnet};
			$stackesubnet = $anchor{$stackesubnet};
			$row[2]       = $stackname;
			$row[1]       = $row[2] . '.uhc.com';
			$row[4]       = deviceType( $row[2] );
			$row[7]       = 'C9300-48P-E';
			$row[12]      = $stacksubnet . '.4';
			$rows         = writeRow( $ws, $rows, \@row, $fmtGreen );
		}
		if ( $anchor{'wlan'} eq 'Y' ) {
			$row[2]  = $anchor{'wlc1_name'};
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'WLC5500';
			$row[12] = $anchor{'svr_subnet_1'} . '.9';    # trivia: wlc is only in XLM sites
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
		}
	} elsif ( $sitetype eq 'L' or $sitetype eq 'X' ) {
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack       = $ct + 1;
			my $stackname   = ( split( /,/, $StackList[$ct] ) )[0];
			my $stacksubnet = 'idf' . $stack . '_ds_1';
			$stacksubnet = $anchor{$stacksubnet};
			$row[2]      = $stackname;
			$row[1]      = $row[2] . '.uhc.com';
			$row[4]      = deviceType( $row[2] );
			$row[7]      = 'C9300-48P-E';
			$row[12]     = $stacksubnet . '.4';
			$rows        = writeRow( $ws, $rows, \@row, $fmtBlue );
		}
		if ( $anchor{'wlan'} eq 'Y' ) {
			$row[2]  = $anchor{'wlc1_name'};
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'WLC5500';
			$row[12] = $anchor{'svr_subnet_1'} . '.9';
			$rows = writeRow( $ws, $rows, \@row, $fmtYellow );
		}
	} elsif ( $sitetype eq 'P' or $sitetype eq 'Q' ) {

		if ( $anchor{'vgcount'} > 0 ) {
			$row[2]  = $anchor{'vgc1_name'};
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'VG204';
			$row[12] = $anchor{'svr_subnet_1'} . '.16';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}

		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack       = $ct + 1;
			my $stacksubnet = 'data_subnet_' . $stack;
			$stacksubnet = $anchor{$stacksubnet};
			my $stackname = ( split( /,/, $StackList[$ct] ) )[0];

			# The only difference between site types is the model
			$row[2]  = $stackname;
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'C9300-48P-E';                            # Q site
			$row[7]  = 'C9300-48P-E' if ( $sitetype eq 'P' );    # P site
			$row[12] = $stacksubnet . '.4';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}
	} elsif ( $sitetype eq 'S' ) {
		for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			my $stack       = $ct + 1;
			my $stackname   = ( split( /,/, $StackList[$ct] ) )[0];
			my $stacksubnet = 'data_subnet_' . $stack;
			$stacksubnet = $anchor{$stacksubnet};

			# The only difference between site types is the model
			my $currentstack = substr( $anchor{'cis1_name'}, -11, );
			$row[2]  = $stackname;
			$row[1]  = $row[2] . '.uhc.com';
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'C9300-48P-E';
			$row[12] = $stacksubnet . '.4';
			$rows = writeRow( $ws, $rows, \@row, $fmtBlue );
		}
	} else {
		prtout("Site Model type '$sitetype' is not valid");
		xit(1);
	}
	$workbook->close();

	return $outputFile;
}

sub writeEquipmentValidation {

	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid ) = @_;
	my $outputFile = $siteid . ' - Equipment Validation Checklist.xls';

	my $workbook = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile")
	  or die "create XLS file '$SVR_ROOTDIR/$OutputDir/$outputFile' failed: $!";

	prtout("Writing Equipment Validation Checklist");
	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75);
	$ws->set_column( 'A:A', 50 );
		#columns hide all for default
	$ws->set_column('B:B', 15, undef,   1); #tad
	$ws->set_column('C:C', 15, undef,   1); #cis1
	$ws->set_column('D:D', 15, undef,   1); #cis2
	$ws->set_column('E:E', 15, undef,   1); #mls1
	$ws->set_column('F:F', 15, undef,   1); #mls2
	$ws->set_column('G:G', 18, undef,   1); #stack
	$ws->set_column('H:H', 15, undef,   1); #wlc/ap
	$ws->set_column('I:I', 15, undef,   1); #wae1
	$ws->set_column('J:J', 15, undef,   1); #wae2
	$ws->set_column('K:K', 15, undef,   1); #vgc1
	$ws->set_column('L:L', 15, undef,   1); #vgc2


	my $fmtBlank = $workbook->add_format( size => 10 );
	my $fmtRedBlank = $workbook->add_format( size => 10 , color => 'red');
	my $fmtGray = $workbook->add_format(
										 size      => 10,
										 bold      => 1,
										 bg_color  => 22,
										 text_wrap => 1
	);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtMergeYellow = $workbook->add_format( size => 13, bold => 1, bg_color => 43, align => 'center');
	my $fmtYellow = $workbook->add_format( size => 13, bold => 1, bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );
	my $fmtNormal = $workbook->add_format( color => 'black', bold => 0 );
	my $fmtBold = $workbook->add_format( bold => 1 );


	# First three rows are the header
	my $rows   = 1;
	my $height = 12;
	$ws->merge_range( 'A1:L1', "Validate $siteid Equipment Checklist", $fmtMergeYellow );

	my $stack;
	my $stackname = ( split( /,/, $StackList[0] ) )[0];
			for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
			$stack = $ct;
			}


	my @header = (
					'',
					$anchor{'tad1_name'},
					$anchor{'cis1_name'},
					$anchor{'cis2_name'},
					$anchor{'mls1_name'},
					$anchor{'mls2_name'},
					$stackname . ' - ' . $stack,
					$anchor{'wlc1_name'} . ' / wap' . $site_code,
					$anchor{'vgc1_name'},
					$anchor{'vgc2_name'},
	);
	$rows = writeRow( $ws, $rows, \@header, $fmtGray );

	my @row = (
			 'Login', '', '', '',
			 '', '', '', '',
			 '', '', '', ''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'Dial in to test the modem', '', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'Ping all other equipment/vlans', 'N/A', '', '',
			 '', '','', '',
			 '', '','', '',
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'IOS - verify its the standard version', '', '', '',
			 '', '', '', 'N/A',
			 '', '', '', ''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show runn - verify the config', '', '', '',
			 '', '', '', '',
			 '', '', '', ''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show cdp neighbor - verify all connections', 'N/A', '', '',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', '', ''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show stand br - verify all the interfaces/vlans', 'N/A', 'N/A', 'N/A',
			 '', '', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show log - look for errors', 'N/A', '', '',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', '', ''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show ip int br - verify all the interfaces', 'N/A', '', '',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', '',
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show ip route - verify all routes are present', 'N/A', '', '',
			 '', '', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show sw detail - verify # of switches, IOS and priority', 'N/A', 'N/A', 'N/A',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show power inline - verify all ports have power', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', '', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show stack-power - verify the priorities', 'N/A', 'N/A', 'N/A',
			 '', '', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '**show inventory - verify the network modules, SFPs, Power', 'N/A', '', '',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '**show env all - verify status of PS, fans, etc.', '', '', '',
			 '', '', '', '',
			 '', '', '', ''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show int trunk - verify the trunking is correct', 'N/A', 'N/A', 'N/A',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '**show diag - verify everything has "passed" diagnostics', 'N/A', '', '',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'verify moh file is there (dir flash:) *only non-SIP sites', 'N/A', '', '',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show vlan - verify the vlans', 'N/A', 'N/A', 'N/A',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show vtp status - should be transparent', 'N/A', 'N/A', 'N/A',
			 '', '', '', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show sysinfo - verify version, system information', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show interface summary', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show wlan summary', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show ap summary***', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 'show ap config general <wap name>***', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 'show disk detail - verify the # of disks and no alarms', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 '', '', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show license - should be Enterprise', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 '', '', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show runn - verify wccp config and that its not enabled', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 '', '', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '', '', '', '',
			 '', '', '', '',
			 '', '', '', ''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '', 'WLC Specifc ', '','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtYellow );

	@row = (
			 'show certificate webauth ***', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 'show network summary (check if webmode is enabled - disable it)', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'show wlan summary', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', '',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '', 'ISR-WAAS Specifc ', '','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtYellow );

	@row = (
			 'show run | I profile  (verify ISR-WAAS-xxx) license count', 'N/A', '', '',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 '', '', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 'dir - verify OVA file', 'N/A', '', '',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 '', '', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 'sh vitrual-service detail', 'N/A', '', '',
			 'N/A', 'N/A', 'N/A', 'N/A',
			 'N/A', 'N/A', 'N/A', 'N/A'
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 '','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '*cis1 notes:','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'Remove bgp network statement so it isnt advertised until cut night','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '*cis2 notes:','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'Remove bgp network statement so it isnt advertised until cut night','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '*stack notes:','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 'Remove EAC (if applicable)','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '**command varies per device type','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = (
			 '** Verify that 2 Power Supplies are present if ordered (cis..01 and 02 routers). ','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 '***The APs are usually not staged so that they connect to the WLC.  ','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 '***So you cannot verify via the WLC, instead ask the staging vendor to take screen shots or somehow give you proof that they staged the APs properly.','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = (
			 '***Check if webauth cert is present: if not create one (then reboot) (needed for https)','','','',
			 '','','','',
			 '','','',''
	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );





	if ( $anchor{'site_type'} eq 'P' ) {
		$ws->set_column('B:B', 15, undef,   0); #tad
		$ws->set_column('C:C', 15, undef,   0); #cis1
		$ws->set_column('G:G', 15, undef,   0); #stack
		$ws->set_column('H:H', 15, undef,   0); #wlc/ap
		$ws->set_column('I:I', 16, undef,   0); #wae1
			if ($anchor{'vgcount'} > 0 ) {
				$ws->set_column('K:K', 15, undef,   0); #vgc1
			}
	}
	if ( $anchor{'site_type'} eq 'S' ){
		$ws->set_column('B:B', 15, undef,   0); #tad
		$ws->set_column('C:C', 15, undef,   0); #cis1
		$ws->set_column('D:D', 15, undef,   0); #cis2
		$ws->set_column('E:E', 15, undef,   0); #stack
		$ws->set_column('H:H', 15, undef,   0); #wlc/ap
		$ws->set_column('I:I', 16, undef,   0); #wae1
		$ws->set_column('J:J', 16, undef,   0); #wae2
			if ($anchor{'vgcount'} > 0 ) {
				$ws->set_column('K:K', 15, undef,   0); #vgc1
			}
	}
	if ( $anchor{'site_type'} eq 'M' ){
		$ws->set_column('B:B', 15, undef,   0); #tad
		$ws->set_column('C:C', 15, undef,   0); #cis1
		$ws->set_column('D:D', 15, undef,   0); #cis2
		$ws->set_column('E:E', 15, undef,   0); #mls1
		$ws->set_column('G:G', 18, undef,   0); #stack
		$ws->set_column('H:H', 15, undef,   0); #wlc/ap
		$ws->set_column('I:I', 16, undef,   0); #wae1
		$ws->set_column('J:J', 16, undef,   0); #wae2
			if ($anchor{'vgcount'} > 0 ) {
				$ws->set_column('K:K', 15, undef,   0); #vgc1
			}
			if ($anchor{'vgcount'} == 2 ) {
				$ws->set_column('L:L', 15, undef,   0); #vgc2
			}
	}
	if (($anchor{'site_type'} eq 'L' ) or ( $anchor{'site_type'} eq 'X' )){
		$ws->set_column('B:B', 15, undef,   0); #tad
		$ws->set_column('C:C', 15, undef,   0); #cis1
		$ws->set_column('D:D', 15, undef,   0); #cis2
		$ws->set_column('E:E', 15, undef,   0); #mls1
		$ws->set_column('F:F', 15, undef,   0); #mls2
		$ws->set_column('G:G', 18, undef,   0); #stack
		$ws->set_column('H:H', 15, undef,   0); #wlc/ap
		$ws->set_column('I:I', 16, undef,   0); #wae1
		$ws->set_column('J:J', 16, undef,   0); #wae2
			if ($anchor{'vgcount'} > 0 ) {
				$ws->set_column('K:K', 15, undef,   0); #vgc1
			}
			if ($anchor{'vgcount'} == 2 ) {
				$ws->set_column('L:L', 15, undef,   0); #vgc2
			}
	}

	$workbook->close();

	return $outputFile;

}

##############################################################################################
#
#
#
##############################################################################################
# Routine to output site data
sub writeSiteRDC {

	my $idfSize = shift;    # default, some site types will pass higher values

	# Site specific info
	my $sitetype = $anchor{'site_type'};
	my $wantype  = $anchor{'pri_circuit_type'};

	$anchor{'cis1_name'}    = 'cis' . $site_code . $anchor{'mdf_flr'} .'a01';
	$anchor{'cis1_ip'}      = $anchor{'loop_subnet'} . ".1";
	$anchor{'cis2_name'}    = 'cis' . $site_code . $anchor{'mdf_flr'} .'a02';
	$anchor{'cis2_ip'}      = $anchor{'loop_subnet'} . ".2";
	$anchor{'mls1_name'}    = 'mls' . $site_code . 'core01';
	$anchor{'mls1_ip'}      = $anchor{'loop_subnet'} . ".33";
	$anchor{'mls2_name'}    = 'mls' . $site_code . 'core02';
	$anchor{'mls2_ip'}      = $anchor{'loop_subnet'} . ".34";
	$anchor{'wae1_name'}    = lc( 'wae' . $site_code . 'core01' );
	$anchor{'wae1_ip'}      = $anchor{'svr_subnet_1'} . ".5";
	$anchor{'wae2_name'}    = lc( 'wae' . $site_code . 'core02' );
	$anchor{'wae2_ip'}      = $anchor{'svr_subnet_1'} . ".6";
	$anchor{'pan1_name'}    = lc( 'pan-' . $site_code . '-01' );
	$anchor{'pan2_name'}    = lc( 'pan-' . $site_code . '-02' );
	$anchor{'mlsdmz1_name'} = lc( 'mls' . $site_code . 'dmz01' );
	$anchor{'mlsdmz2_name'} = lc( 'mls' . $site_code . 'dmz02' );
	$anchor{'bip1_name'}    = lc( 'bip' . $site_code . 'prd01' );
	$anchor{'bip2_name'}    = lc( 'bip' . $site_code . 'prd02' );
	$anchor{'bipapp1_name'} = lc( 'bip' . $site_code . 'app01' );
	$anchor{'bipapp2_name'} = lc( 'bip' . $site_code . 'app02' );
	$anchor{'rtr1_name'}    = lc( 'rtr' . $site_code . 'gre01' );
	$anchor{'rtr2_name'}    = lc( 'rtr' . $site_code . 'gre02' );

	# core DCF5
	$anchor{'ltm1_name'}    = lc( 'ltm' . $site_code . 'prd01' );
	$anchor{'ltm2_name'}    = lc( 'ltm' . $site_code . 'prd02' );
	$anchor{'mlsdis1_name'} = lc( 'mls' . $site_code . 'dis01' );
	$anchor{'mlsdis2_name'} = lc( 'mls' . $site_code . 'dis02' );

	# tools
	$anchor{'das_name'}    = lc( 'das' . $site_code . 'prd01' );
	$anchor{'mlsinf_name'} = lc( 'mon' . $site_code . 'inf01' );    # change to mon
	$anchor{'mlscor_name'} = lc( 'mls' . $site_code . 'cor' );
	$anchor{'tad1_name'}   = 'tad' . $site_code . $anchor{'mdf_flr'} .'a01';

	# wgs 1-30
	for ( my $d = 1 ; $d <= 30 ; $d++ ) {
		$anchor{ 'wgs' . $d . '_name' } =
		  lc( sprintf( "wgs%spri%02d", $site_code, $d ) );
	}
	my %sitemask = ( 'RDC', '255.255.254.0', );
	$anchor{'site_mask'} = $sitemask{$sitetype};
	if ( $anchor{'fw'} eq 'Y' ) {
		my $null;
		( $null, $null, my $octet3 ) = split( /\./, $anchor{'site_mask'} );
		( my $fwoctet1, my $fwoctet2, my $fwoctet3 ) =
		  split( /\./, $anchor{'loop_subnet'} );
		my $fwoctet3_new = $fwoctet3 + ( ( 254 - $octet3 ) / 2 );
		my $octet3_last  = $fwoctet3 + ( ( 254 - $octet3 ) / 2 );
		$anchor{'fwi_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $fwoctet3_new;
		if ( $sitetype eq 'X' ) {
			my $octetdmz = $octet3_last - 1;
			$anchor{'dmz_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $octetdmz;
		} elsif ( $sitetype eq 'L' ) {
			my $octetdmz = $octet3_last - 2;
			$anchor{'dmz_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $octetdmz;
		} elsif ( $sitetype eq 'M' ) {
			my $octetdmz = $octet3_last - 2;
			$anchor{'dmz_subnet'} = $fwoctet1 . '.' . $fwoctet2 . '.' . $octetdmz;
		}
	}
	if( $anchor{'proj_type'} eq 'build'){
		# MLS and Stack configs
		prtout("");
		prtout("Writing MLS Configurations");
		writeTemplate( 'Model-RDC/mls1.txt', $anchor{'mls1_name'} . '.txt' );
		writeTemplate( 'Model-RDC/mls2.txt', $anchor{'mls2_name'} . '.txt' );

		# tad configs
		prtout("Writing TAD Configurations");
		# set tad port names
		$anchor{'tad_port_1'}  = $anchor{'cis1_name'};
		$anchor{'tad_port_2'}  = $anchor{'cis2_name'};
		$anchor{'tad_port_3'}  = $anchor{'mls1_name'};
		$anchor{'tad_port_4'}  = $anchor{'mls2_name'};
		$anchor{'tad_port_5'}  = $anchor{'pan1_name'};
		$anchor{'tad_port_6'}  = $anchor{'pan2_name'};
		$anchor{'tad_port_7'}  = 'Optional';
		$anchor{'tad_port_8'}  = 'Optional';
		$anchor{'tad_port_9'}  = 'Optional';
		$anchor{'tad_port_10'} = 'Optional';
		$anchor{'tad_port_11'} = 'Optional';
		$anchor{'tad_port_12'} = 'Optional';
		$anchor{'tad_port_13'} = 'Optional';
		$anchor{'tad_port_14'} = 'Optional';
		$anchor{'tad_port_15'} = 'Optional';
		$anchor{'tad_port_16'} = 'Optional';
		$anchor{'tad_hosts'} = smbReadFile("Modules/rdc_tad_hosts.txt");
		writeTemplate( "Generic-OOB/tad_g526.txt", 'tad' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01_G526.txt' );
		writeTemplate( "Generic-OOB/g526.txt", 'lte' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01.txt' );

		# For whatever reason the 'biotab' state needs to be passed to the clean version, then added to again
		setRDCAnchors();
		writeVisioRDC( $sitetype, $idfSize, 'Normal' );
		writeVisioRDC( $sitetype, $idfSize, 'Clean' ); #Added clean version of DnE


		# writeVisioRDC( $sitetype, $stacklimit{$sitetype}, 'Clean' );
		prtout("Writing IP Summary XLS");
		my $ipSummaryFile;
		if ( $anchor{'uhgdivision'} eq 'UHG' ) {
			$ipSummaryFile = writeIPSummaryRDC( $anchor{'site_code'}, $sitetype, $idfSize );

			#   do we need a ci summary ?
			writeCISummaryRDC( $anchor{'site_code'}, $sitetype, $idfSize );
		}
		#Force SDWAN anchor to YES
		$anchor{'SDWAN'} = 'Y';
		writeSDWANcsv($anchor{'site_code'},$anchor{'cis1_name'},$anchor{'cis2_name'},$anchor{'router_seltype'}, $anchor{'int_type'}, $anchor{'int_type_r1'}, $anchor{'int_type_r2'}, $anchor{'transport'});
		unlink("$ROOTDIR/Files/$ipSummaryFile");
	}
	elsif( $anchor{'proj_type'} eq 'proj-sdwan'){
		#Force SDWAN anchor to YES
		$anchor{'SDWAN'} = 'Y';
		writeSDWANcsv($anchor{'site_code'},$anchor{'cis1_name'},$anchor{'cis2_name'},$anchor{'router_seltype'}, $anchor{'int_type'}, $anchor{'int_type_r1'}, $anchor{'int_type_r2'}, $anchor{'transport'});
	}

	my $zipfile = compress();

	prtout( "Configurator Output Generation Complete.<br/>",
		"<a HREF='/tmp/$OutputDir.zip' >D&E and IP Summary can be found here</a>" );
}
##############################################################################################
#
#
#
##############################################################################################
sub writeVisioRDC {
	return unless ($anchor{'proj_type'} eq 'build');
	$anchor{'street'} =~s/&/&amp;/g;
	( my $sitetype, my $idfSize, my $vistype ) = @_;
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-rdc-template.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-rdc-template-clean.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';
	}
	my $backmdftab   = 5;
	my $mdfrdctab      = 11;
	my $backbiotab   = 16;
	my $biographytab = 17;
	my $fwovertab    = 23;
	my $backsectab   = 24;
	my $secl3tab     = 25;
	my $tadtab       = 27;
	my $untrusttab   = 37;

	#		PGID: '0' NameU: 'Background-DC' Name: 'Background-DC'
	#		PGID: '4' NameU: 'Background-DC1' Name: 'Background-IDF'
	#		PGID: '5' NameU: 'Background' Name: 'Background-MDF'
	#		PGID: '10' NameU: 'Background-Stacks' Name: 'Background-Stacks'
	#		PGID: '12' NameU: 'Background-DC_Traditional' Name: 'Background-DC_Traditional'
	#		PGID: '14' NameU: 'Background-IPT' Name: 'Background-IPT'
	#		PGID: '16' NameU: 'Background-Bio' Name: 'Background-Bio'
	#		PGID: '24' NameU: 'Background-Security' Name: 'Background-Security'
	#		PGID: '30' NameU: 'Background-WLAN' Name: 'Background-WLAN'
	#		PGID: '17' NameU: 'Biography' Name: 'Biography'
	#		PGID: '54' NameU: 'Core DC' Name: 'Core DC F5 LB'
	#		PGID: '11' NameU: 'MDF/ Template' Name: 'Template-MDF C'
	#		PGID: '2' NameU: 'Data Center' Name: 'Template-IDF 1 C'
	#		PGID: '6' NameU: 'Data Center 2' Name: 'Template-IDF 2 C'
	#		PGID: '7' NameU: 'Data Center 3' Name: 'Template-IDF 3 C'
	#		PGID: '27' NameU: 'TAD' Name: 'TAD'
	#		PGID: '25' NameU: 'Security Layer3' Name: 'M&amp;A Untrusted'
	#		PGID: '37' NameU: 'Untrusted/DMZ' Name: 'Untrusted/DMZ'
	#		PGID: '1' NameU: 'DMZ Flow' Name: 'DMZ Flow'
	#		PGID: '23' NameU: 'ASA Overview' Name: 'FW Overview'
	#		PGID: '33' NameU: 'Stack Overview' Name: 'Stack Overview'
	#		PGID: '34' NameU: 'WAE 574-7371 Overview' Name: 'WAE Overview'
	#		PGID: '13' NameU: 'Tools' Name: 'Tools'

	prtout("Opening Visio Template for Processing");
	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );

	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);

	# Upload tabs
	my %tabs = VisioReadTabs( \$fill );

	# Add Bio Page
	my $biotab = $tabs{$biographytab};

	#Add MDF and TAD tab
	my $mdftab = $tabs{$mdfrdctab} . $tabs{$tadtab};

	#Security tabs
	my $sectab = $tabs{$secl3tab} . $tabs{$untrusttab} . $tabs{$fwovertab};

	#Add selected Router type
	VisioControlLayer( 'ASR',                  0, \$mdftab );
	VisioControlLayer( '4461',                 0, \$mdftab );
	VisioControlLayer( 'C8300-1N1S',           0, \$mdftab );
	VisioControlLayer( 'C8300-2N2S',           0, \$mdftab );
	VisioControlLayer( 'C8500-12X4QC',           0, \$mdftab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$mdftab );

	#SDWAN
	#added selection for 4-transport and DUAL-DIA
	VisioControlLayer( '2_transports', 0, \$mdftab )
	if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} eq 'yes');

	VisioControlLayer( '3_transports', 0, \$mdftab )
	if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} eq 'yes');

	VisioControlLayer( '2_transpo_new', 0, \$mdftab )
	if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} ne 'yes');

	VisioControlLayer( '3_transpo_new', 0, \$mdftab )
	if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} ne 'yes');

	VisioControlLayer( '4_transpo_new', 0, \$mdftab )
	if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 4 or $anchor{'tloc'} ne 'yes');

	#background tabs
	my $backtab = $tabs{$backbiotab} . $tabs{$backmdftab} . $tabs{$backsectab};

	# Search and replace variables
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
		$sectab =~ s/\!$key\!/$anchor{$key}/g;
		$backtab =~ s/\!$key\!/$anchor{$key}/g;
	}
	my $generatedTabs = $biotab . $mdftab . $sectab . $backtab;

	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");

	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

}
############################################################################
#
#
############################################################################
sub writeIPSummaryRDC {
	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid, my $sitetype, my $stacklimit ) = @_;
	my $outputFile = $siteid . '-IP-Summary-Chart.xls';
	my $workbook   = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile")
	  or die "create XLS file '$ROOTDIR/Files/$outputFile' failed: $!";
	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75);
	$ws->set_column( 'A:A', 14.5 );
	$ws->set_column( 'B:B', 16.5 );
	$ws->set_column( 'C:C', 16.5 );
	$ws->set_column( 'D:D', 16.5 );
	$ws->set_column( 'E:E', 32 );
	$ws->set_column( 'F:F', 23 );
	my $fmtBlank = $workbook->add_format( size => 12, bold => 1 );
	my $fmtGray = $workbook->add_format(
										 size      => 9,
										 bold      => 1,
										 bg_color  => 22,
										 text_wrap => 1
	);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtYellow = $workbook->add_format( bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );

	# First three rows are the header
	my $rows   = 2;
	my $height = 12;
	$ws->merge_range( 'A1:H2', "$siteid - IP Summary Chart", $fmtBlank );
	my @header = (
				   'Status',
				   'IDF',
				   'IP Subnet',
				   'IP Address',
				   'Device Name',
				   'Device Description',
				   'DHCP',
				   'DNS',
	);
	$ws->write_row( $rows, 0, \@header, $fmtGray );
	$rows++;

	# CIS values
	my %cistypes = ( 'ASR', 'ASR1001-X', '4451', 'C4451', '4461', 'C4461', 'C8200-1N-4T', 'C8200-1N-4T', 'C8300-1N1S', 'C8300-1N1S', 'C8300-2N2S', 'C8300-2N2S', 'C8500-12X4QC', 'C8500-12X4QC' );
	my $cistype = $cistypes{$anchor{'router_type'}};

	# MLS values
	my $mlstype = 'N93180';

	# Decide whether Voice VLANs should be entered as DHCP scopes or not
	my $vdhcp = 'No';
	$vdhcp = 'Yes*' if ( $anchor{'ipt'} eq 'Y' );

	# AT&T CER-PER
	my $wan_subnet;
	if ( $anchor{'pri_wan_ip_cer'} ne '' ) {
		( my $octet1, my $octet2, my $octet3, my $octet4 ) =
		  split( /\./, $anchor{'pri_wan_ip_cer'} );
		$octet4--;
		$wan_subnet = join( '.', $octet1, $octet2, $octet3, $octet4 );
		}

	# Trimming out "anchor" to shorten horizontal length
	my $pri_wan_ip_per = $anchor{'pri_wan_ip_per'};
	my $pri_wan_ip_cer = $anchor{'pri_wan_ip_cer'};
	my $att_upl_int = $anchor{'att_upl_int_dns'};
	my $cis_mls_int = $anchor{'cis_mls_int'};
	my $cis_mls_int2 = $anchor{'cis_mls_int2'};
	my $tad1_ip = $anchor{'tad1_ip'};
	my $loop_subnet = $anchor{'loop_subnet'};
	my $svr_subnet_1 = $anchor{'svr_subnet_1'};
	my $tad1_subnet = $anchor{'tad1_subnet'};
	my $tad1_name = $anchor{'tad1_name'};
	my $mls1_name = $anchor{'mls1_name'};
	my $mls2_name = $anchor{'mls2_name'};
	my $cis1_name = $anchor{'cis1_name'};
	my $cis2_name = $anchor{'cis2_name'};
	my $r1_vlan = $anchor{'r1_vlan'};
	my $tloc_int = $anchor{'tloc_int'};
	my $tloc_int2 = $anchor{'tloc_int2'};

	# Start writing output
	my @row = ( 'New', 'mdf', '', '', '', '', 'No', 'Yes' );
	my @ipsum = (
		[ $tad1_subnet, 		   $tad1_ip,  		   $tad1_name, 				  'SLC8000', 			'RDC' ],#ip ==  0
		[ "$loop_subnet.1/32",     "$loop_subnet.1",   $cis1_name,				  $cistype , 			'RDC'      ],#ip ==  1
		[ "$loop_subnet.2/32",     "$loop_subnet.2",   $cis2_name, 				  $cistype ,  			'RDC'    ],#ip ==  2
		[ "$loop_subnet.33/32",    "$loop_subnet.33",   $mls1_name, 				  $mlstype, 			'RDC'    ],#ip ==  3
		[ "$loop_subnet.34/32",    "$loop_subnet.34",   $mls2_name, 				  $mlstype,  			'RDC'	   ],#ip ==  4
		[ "$wan_subnet/30",		   $pri_wan_ip_cer,    "$cis1_name-$att_upl_int$r1_vlan", 'WAN Circuit', 		'RDC' ],#ip == 5
		[ "$wan_subnet/30",		   $pri_wan_ip_per,	   "per-$cis1_name",		  'WAN Circuit', 		'RDC' ],#ip == 6
		[ "$loop_subnet.128/30",    "$loop_subnet.129",  "$cis1_name-$cis_mls_int", 	  'cis1 to mls1', 		'RDC' 	   ],#ip == 9
		[ "$loop_subnet.128/30",    "$loop_subnet.130",  "$mls1_name-eth1-1", 	  'mls1 to cis1', 		'RDC'   	   ],#ip == 10
		[ "$loop_subnet.132/30",    "$loop_subnet.133",  "$cis1_name-$cis_mls_int2", 	  'cis1 to mls2', 		'RDC'       ],#ip == 11
		[ "$loop_subnet.132/30",    "$loop_subnet.134",  "$mls2_name-eth1-1", 	  'mls2 to cis1', 		'RDC' 	   ],#ip == 12
		[ "$loop_subnet.136/30",    "$loop_subnet.137",  "$cis2_name-$cis_mls_int2", 	  'cis2 to mls1', 		'RDC' 	   ],#ip == 13
		[ "$loop_subnet.136/30",    "$loop_subnet.138",  "$mls1_name-eth1-2", 	  'mls1 to cis2', 		'RDC'	   ],#ip == 14
		[ "$loop_subnet.140/30",    "$loop_subnet.141",  "$cis2_name-$cis_mls_int", 	  'cis2 to mls2', 		'RDC'	   ],#ip == 15
		[ "$loop_subnet.140/30",    "$loop_subnet.142",  "$mls2_name-eth1-2",   	  'mls2 to cis2', 		'RDC'	   ],#ip == 16
		[ "$loop_subnet.144/30",    "$loop_subnet.145",  "$mls1_name-vlan-100",	  'mls1 to mls2', 		'RDC'	   ],#ip == 17
		[ "$loop_subnet.144/30",    "$loop_subnet.146",  "$mls2_name-vlan-100", 	  'mls2 to mls1',		'RDC'	   ],#ip == 18
		[ $tad1_subnet, 		   "$loop_subnet.162",  "$mls1_name-vlan-901", 	  'mls1 to tad', 		'RDC'	   ],#ip == 19
		[ "$loop_subnet.80/28",    "$loop_subnet.81",  "$mls1_name-gwy-vlan-99", 	  'Firewall Inside Vlan GWY',  'RDC'       ],#ip == 20
		[ "$loop_subnet.80/28",    "$loop_subnet.82",  "$mls1_name-vlan-99", 	  'Firewall Inside mls1',  'RDC'       ],#ip == 21
		[ "$loop_subnet.80/28",    "$loop_subnet.83",  "$mls2_name-vlan-99", 	  'Firewall Inside mls2',  'RDC'       ],#ip == 22
		[ "$svr_subnet_1.0/24",    "$svr_subnet_1.1",  "$mls1_name-gwy-vlan-101", 	  'Server Vlan GWY',  'RDC'       ],#ip == 23
		[ "$svr_subnet_1.0/24",    "$svr_subnet_1.2",  "$mls1_name-vlan-101", 	  'Server Vlan mls1',  'RDC'       ],#ip == 24
		[ "$svr_subnet_1.0/24",    "$svr_subnet_1.3",  "$mls2_name-vlan-101", 	  'Server Vlan mls2',  'RDC'       ],#ip == 25
	);
	for ( my $ip = 0; $ip <= $#ipsum; $ip++ ) { # $ip will be used as parent array's index
		$row[2] = $ipsum[$ip][0];
		$row[3] = $ipsum[$ip][1];
		$row[4] = $ipsum[$ip][2];
		$row[5] = $ipsum[$ip][3];
		$row[1] = 'mdf';
		$row[1] = 'Vlan101/mdf' if ( $row[4] =~ /vlan-101/ );
		$row[1] = 'Vlan100/mdf' if ( $row[4] =~ /vlan-100/ );
		$row[1] = 'Vlan901/mdf' if ( $row[4] =~ /vlan-901/ );
		$row[1] = 'Vlan99/mdf' if ( $row[4] =~ /vlan-99/ );
		$row[6] = 'No';
		my $fmtColor = $fmtPurple;
		$fmtColor = $fmtBlue if ( $row[4] =~ /-/ );

		if ( $ipsum[$ip][4] =~ /$sitetype/ and $row[4] =~ /(tad|cis|mls|stk|vgc|wlc)/ ) {
			$rows = writeRow( $ws, $rows, \@row, $fmtColor );
		}

		if ( $ip == 6  and $anchor{'SDWAN'} eq 'Y' ) {
		$anchor{'tloc_int'} =~ s/\//-/g;
		my $tloc_int_dns = lc($anchor{'tloc_int'});
		$anchor{'tloc_int2'} =~ s/\//-/g;
		my $tloc_int2_dns = lc($anchor{'tloc_int2'});
		$anchor{'tloc_int3'} =~ s/\//-/g;
		my $tloc_int3_dns = lc($anchor{'tloc_int3'});
		$anchor{'tloc_int4'} =~ s/\//-/g;
		my $tloc_int4_dns = lc($anchor{'tloc_int4'});

		# SDWAN TLOC Extensions
		if( $anchor{'tloc'} eq 'yes' ){
				if ( $anchor{'transport'} == 2){
				# Router 1
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.48/30";
				$row[3] = "$loop_subnet.49";
				$row[4] = "$cis1_name-$tloc_int_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.50";
				$row[4] = "$cis2_name-$tloc_int_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2
				$row[2] = "$loop_subnet.52/30";
				$row[3] = "$loop_subnet.53";
				$row[4] = "$cis1_name-$tloc_int2_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.54";
				$row[4] = "$cis2_name-$tloc_int2_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
				if ( $anchor{'transport'} == 3){
				# Router 1
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.48/30";
				$row[3] = "$loop_subnet.49";
				$row[4] = "$cis1_name-$tloc_int_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.50";
				$row[4] = "$cis2_name-$tloc_int_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2
				$row[2] = "$loop_subnet.52/30";
				$row[3] = "$loop_subnet.53";
				$row[4] = "$cis1_name-$tloc_int2_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.54";
				$row[4] = "$cis2_name-$tloc_int2_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 1 - MPLS,Private1,Private2 TLOC-interface
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.XX/30";
				$row[3] = "$loop_subnet.XX";
				$row[4] = "$cis1_name-$tloc_int3_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.XX";
				$row[4] = "$cis2_name-$tloc_int3_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
				if ( $anchor{'transport'} == 4){
				# Router 1
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.48/30";
				$row[3] = "$loop_subnet.49";
				$row[4] = "$cis1_name-$tloc_int_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.50";
				$row[4] = "$cis2_name-$tloc_int_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2
				$row[2] = "$loop_subnet.52/30";
				$row[3] = "$loop_subnet.53";
				$row[4] = "$cis1_name-$tloc_int2_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.54";
				$row[4] = "$cis2_name-$tloc_int2_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 1 - MPLS,Private1,Private2 TLOC-interface
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.XX/30";
				$row[3] = "$loop_subnet.XX";
				$row[4] = "$cis1_name-$tloc_int3_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.XX";
				$row[4] = "$cis2_name-$tloc_int3_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2 - MPLS,Private1,Private2 TLOC-interface
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.YY/30";
				$row[3] = "$loop_subnet.YY";
				$row[4] = "$cis1_name-$tloc_int4_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.YY";
				$row[4] = "$cis2_name-$tloc_int4_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
		}
		else{
		# Router 1
		$row[1] = 'mdf';
		$row[2] = "$loop_subnet.48/30";
		$row[3] = "$loop_subnet.49";
		$row[4] = "$cis1_name-$tloc_int_dns";
		$row[5] = 'cis1 TLOC extension';
		$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
		# Router 2
		$row[3] = "$loop_subnet.50";
		if ( $anchor{'transport'} == 2){
			$row[4] = "$cis2_name-$tloc_int_dns";
		}
		else{
			$row[4] = "$cis2_name-$tloc_int2_dns";
		}
		$row[5] = 'cis2 TLOC extension';
		$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

		# 2 transports only
		if ( $anchor{'transport'} == 2) {
			$row[2] = "$loop_subnet.52/30";
			$row[3] = "$loop_subnet.53";
			#$row[4] = "$cis1_name-g0-0-2-20";
			$row[4] = "$cis1_name-$tloc_int2_dns";
			$row[5] = 'cis1 TLOC extension';
			# prtout("IP summary Row4: $row[4]");
			$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
			$row[3] = "$loop_subnet.54";
			#$row[4] = "$cis2_name-g0-0-2-20";
			$row[4] = "$cis2_name-$tloc_int2_dns";
			$row[5] = 'cis2 TLOC extension';
			# prtout("IP summary Row4: $row[4]");
			$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
			}
		}
	}
	}

	$rows++;    # leave a blank line before the notes and legend
	@row = ( 'Key', '', 'if applicable---->', '* - Requires option 150 TFTP Servers' );
	$rows = writeRow( $ws, $rows, \@row );

	# Formats are different for the cells in this row, so write them directly instead of passing to the sub
	$ws->write( $rows, 0, 'Loopback/Mgmt', $fmtPurple );
	$ws->write( $rows, 3, '* - Please add option 150 for the voice DHCP scope per IPT Engineering Team Standards' );
	$rows++;
	$ws->write( $rows, 0, 'Router/MLS', $fmtBlue );
	$ws->write( $rows, 3, '** - Wireless LAN DHCP scopes should be configured with a 2-hour lease time' );
	$rows++;

	return $outputFile;
}

sub writeCISummaryRDC {
	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid, my $sitetype, my $stacklimit ) = @_;
	my $outputFile = $siteid . '-CI-Devices.xls';
	my $workbook   = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile");
	prtout("Writing CI Summary XLS");
	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75);
	$ws->set_column( 'A:A', 14.5 );
	$ws->set_column( 'B:B', 30 );
	$ws->set_column( 'C:C', 17 );
	$ws->set_column( 'D:D', 18 );
	$ws->set_column( 'E:E', 25 );
	$ws->set_column( 'F:F', 16.5 );
	$ws->set_column( 'G:G', 16.5 );
	$ws->set_column( 'H:H', 18 );
	$ws->set_column( 'I:I', 23 );
	$ws->set_column( 'J:J', 13 );
	$ws->set_column( 'K:K', 30 );
	$ws->set_column( 'L:L', 12 );
	$ws->set_column( 'M:M', 12 );
	my $fmtBlank = $workbook->add_format( size => 12, bold => 1 );
	my $fmtGray = $workbook->add_format(
										 size      => 9,
										 bold      => 1,
										 bg_color  => 22,
										 text_wrap => 1
	);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtYellow = $workbook->add_format( bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );

	# First three rows are the header
	my $rows   = 2;
	my $height = 12;
	$ws->merge_range( 'A1:O2', "$siteid - CI Summary Chart", $fmtBlank );
	my @header = (
				   'If Update to existing CI, provide HPSM CI ID',
				   'SYSTEM ID/FQDN (OVO ID)',
				   'FORMAL NAME',
				   'COMMON NAME',
				   'CI TYPE',
				   'ENVIRONMENT',
				   'SERIAL NUMBER',
				   'MODEL',
				   'BRAND',
				   'LOCATION',
				   'SUPPORTING WORKGROUP',
				   'REMARKS',
				   'IP ADDRESS',
				   'IP ADDRESS (additional)'
	);
	$rows = writeRow( $ws, $rows, \@header, $fmtGray );

	# CIS values
	my %cistypes = ( 'ASR', 'ASR1001-X', '4451', 'C4451', '4461', 'C4461', 'C8200-1N-4T', 'C8200-1N-4T', 'C8300-1N1S', 'C8300-1N1S', 'C8300-2N2S', 'C8300-2N2S', 'C8500-12X4QC', 'C8500-12X4QC' );
	my $cistype = $cistypes{$anchor{'router_type'}};

	# MLS values
	my $mlstype = 'N93180';

	my ( $mlsupl, $mlsupl1, $mlsupl2, $mlsinter );

	# Supporting workgroup is always the same
	my $workgroup = 'NHS_SLO_Remote'; # Per Sonja 04/30/2020

	# Write the output
	my $domain = '.uhc.com';
	my @row = ( '', '', '', '', '', 'Production', '', '', 'Cisco', $siteid, $workgroup, '', '', '' );
	$row[2]  = $anchor{'cis1_name'};
	$row[1]  = $row[2] . $domain;
	$row[4]  = deviceType( $row[2] );
	$row[7]  = $cistype;
	$row[12] = $anchor{'loop_subnet'} . '.1';
	$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	$row[2]  = $anchor{'cis2_name'};
	$row[1]  = $row[2] . $domain;
	$row[4]  = deviceType( $row[2] );
	$row[7]  = $cistype;
	$row[12] = $anchor{'loop_subnet'} . '.2';
	$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	$row[2]  = $anchor{'mls1_name'};
	$row[1]  = $row[2] . $domain;
	$row[4]  = deviceType( $row[2] );
	$row[7]  = $mlstype;
	$row[12] = $anchor{'loop_subnet'} . '.33';
	$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	$row[2]  = $anchor{'mls2_name'};
	$row[1]  = $row[2] . $domain;
	$row[4]  = deviceType( $row[2] );
	$row[7]  = $mlstype;
	$row[12] = $anchor{'loop_subnet'} . '.34';
	$rows = writeRow( $ws, $rows, \@row, $fmtPurple );
	$row[2]  = $anchor{'tad1_name'};
	$row[1]  = $row[2] . $domain;
	$row[4]  = deviceType( $row[2] );
	$row[7]  = 'SLC8000';
	$row[8]  = 'Lantronix';
	$row[12] = $anchor{'loop_subnet'} . '.161';
	$rows = writeRow( $ws, $rows, \@row, $fmtPurple );

	$workbook->close();
	prtout("Copying CI summary file to finished directory");
	return $outputFile;
}

sub setRDCAnchors {
	my @octs    = split /\./, ( $anchor{'loop_subnet'} . ".0" );
	my $oct24   = $octs[2];
	my $vlan600 = 600;
	for our $subnet ( 0 ... 63 ) {
		$octs[3] = 0;    # reset the last octet to 0
		if ( $subnet == 0 ) {
			my $label = "vlan_99_sn";
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

		} elsif ( $subnet == 1 ) {
			my $label = "vlan_101_sn";
			$octs[2] += 1;
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

		} elsif ( $subnet == 2 ) {
			my $label = "Optum Management/InterMLS Network space";
			$octs[2] += 1;

		} elsif ( ( $subnet >= 3 ) and ( $subnet <= 55 ) ) {
			if ( ( $subnet == 8 ) or ( $subnet == 12 ) ) {
				$octs[2] += 1;
				my $label = sprintf( "vlan_%s_sn", $vlan600++ );
				$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

				$octs[2] += 3;
			} elsif ( ( $subnet > 8 ) and ( $subnet < 16 ) ) {
				next;
			} else {
				$octs[2] += 1;
				my $label = sprintf( "vlan_%s_sn", $vlan600++ );
				$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

			}
		} elsif ( $subnet == 58 ) {
			$octs[2] += 1;
			my $label = qq(vlan_500_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

			$octs[3] += 128;
			$label = qq(vlan_501_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

		} elsif ( $subnet == 59 ) {
			$octs[2] += 1;
			my $label = qq(vlan_502_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

		} elsif ( $subnet == 60 ) {
			$octs[2] += 1;
			my $label = qq(vlan_503_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

		} elsif ( $subnet == 61 ) {
			$octs[2] += 1;
			my $label = qq(vlan_504_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

			$octs[3] += 128;
			$label = qq(vlan_505_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

		} elsif ( $subnet == 62 ) {
			$octs[2] += 1;
			my $label = qq(vlan_506_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

			$octs[3] += 128;
			$label = qq(vlan_507_sn);
			$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

		} elsif ( $subnet == 63 ) {
			$octs[2] += 1;
			my $label = "";
			my $vlan  = 0;
			for ( my $prt = 0 ; $prt < 15 ; $prt++ ) {
				if ( $prt == 0 ) {
					$label = qq(vlan_10_sn);
				} else {
					if ( ( $prt == 13 ) or ( $prt == 14 ) ) {
						$octs[3] += 16;
						$vlan += 16;
						$label = 'vlan_' . $vlan . '_sn';
					} else {
						$octs[3] += 8;
						$vlan += 8;
					}
					$label = 'vlan_' . $vlan . '_sn';
				}
				$anchor{$label} = join( '.', @octs[ 0 ... 2 ] );

			}
		} else {
			$octs[2] += 1;
		}
	}
}

sub writeSDWANcsv {
	prtout( "Writing SDWAN CSV.....");

	( my ($siteid, $r1 , $r2, $rt_type,$int_type,$int_typer1,$int_typer2,$tport) ) = @_;
	my $outputFile;
	my @PST = ('WA','OR','NV','CA');
	my @MST = ('MT','ID','WY','UT','CO','AZ','NM');
	my @CST = ('ND','SD','MN','WI','NE','IA','IL','KS','MO','KY','OK','AR','TN','TX','LA','MS','AL');
	my @EST = ('MI','IN','GA','OH','WV','PA','VA','NC','SC','FL','VT','NH','ME','DC','MD','DE','NJ','CT','MA','RI','NY');
	my $site = uc(substr($r1, 3,2));
	my $region;


	if($site ~~ @PST){
		$region = "PST";
	}
	elsif($site ~~ @MST){
		$region = "MST";
	}
	elsif($site ~~ @CST){
		$region = "CST";
	}
	elsif($site ~~ @EST){
		$region = "EST";
	}
	prtout( "Region: $region");

	if($rt_type == "4451"){
			$rt_type = '4451-X';
	}
	my $rtr_type = '';
	if (($rt_type eq '4331') or ($rt_type eq '4451-X') or ($rt_type eq '4461')){
		$rtr_type = "ISR" . $rt_type;
	}else{
		$rtr_type = $rt_type;
	}

	# my $rtr_type = "ISR" . $rt_type;
	$anchor{'street'} =~ s/,|\.//g;
	$anchor{'street'} =~ s/ /_/g; #replace spcaes with '_';

	my $MPLS_provider = '';
	my $MPLS_provider2 = '';
	if (($anchor{'r1_provider'}) eq 'ATT'){
		$MPLS_provider = 'ATT';
	}
	elsif (($anchor{'r1_provider'} eq 'VZ') and ($anchor{'pri_vlan'} > 0)){
		$MPLS_provider = 'VZ';
	}
	elsif (($anchor{'r1_provider'} eq 'VZ') and ($anchor{'pri_vlan'} == 0)){
		$MPLS_provider = "VZ_un";
	}
	elsif (($anchor{'r1_provider'} eq 'LUMEN') and ($anchor{'pri_vlan'} > 0)){
		$MPLS_provider = 'Lumen';
	}
	elsif (($anchor{'r1_provider'} eq 'LUMEN') and ($anchor{'pri_vlan'} == 0)){
		$MPLS_provider = "Lumen_un";
	}
	if (($anchor{'r2_provider'}) eq 'ATT'){
		$MPLS_provider2 = 'ATT';
	}
	elsif (($anchor{'r2_provider'} eq 'VZ') and ($anchor{'r2vlan'} > 0)){
		$MPLS_provider2 = 'VZ';
	}
	elsif (($anchor{'r2_provider'} eq 'VZ') and ($anchor{'r2vlan'} == 0)){
		$MPLS_provider2 = "VZ_un";
	}
	elsif (($anchor{'r2_provider'} eq 'LUMEN') and ($anchor{'r2vlan'} > 0)){
		$MPLS_provider2 = 'Lumen';
	}
	elsif (($anchor{'r2_provider'} eq 'LUMEN') and ($anchor{'r2vlan'} == 0)){
		$MPLS_provider2 = "Lumen_un";
	}

	my $pico_folder = "SDWAN_Temp_Legacy/Pico/SDWAN_Temp_Summarized/";
	my $small_folder = "SDWAN_Temp_Legacy/Small/SDWAN_Temp_Summarized/";
	my $med_folder = "SDWAN_Temp_Legacy/Medium/SDWAN_Temp_Summarized/";
	my $large_folder = "SDWAN_Temp_Legacy/Large/SDWAN_Temp_Summarized/";
	my $RDC_folder = "SDWAN_Temp_Legacy/RDC/Consolidated/";
	my $stg_folder = "SDWAN_Temp_Legacy/Staging/";

	if ( $anchor{'site_type'} eq 'P' ) {
			$outputFile = writeTemplate( $pico_folder . "Trust_RS_Consolidated_" . $rtr_type ."_Pico_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_RS_'. $region . '_' . $rtr_type .'_Pico_' . $int_type . '_dtmpl.csv');
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
			}
	}
	if ( $anchor{'site_type'} eq 'S' ) {
			$outputFile = writeTemplate( $small_folder . "Trust_RS_Consolidated_" . $rtr_type ."_Small_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_'. $region . '_' . $rtr_type .'_Small_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			$outputFile = writeTemplate( $small_folder . "Trust_RS_Consolidated_" . $rtr_type ."_Small_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region .'_' . $rtr_type . ' _Small_Sec_' . $tport . '-Transport_dtmpl.csv' );
			if ($int_type eq 'static'){
				if ($tport == 2){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}else{
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}if ($int_type eq 'dhcp'){
				if ($tport == 2){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_DHCP_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				}
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}
	if ( $anchor{'site_type'} eq 'M' ) {

			if( $anchor{'tloc'} eq 'yes' ){
			if ($tport == 2){
			  $outputFile = writeTemplate( $med_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Small_Pri_" . $int_type . "_" . $tport . "-Transport_INET-EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Small_Pri_' . $int_type . '_' . $tport . '-Transport_INET-EXT_dtmpl.csv');
			  $outputFile = writeTemplate( $med_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Small_Sec_" . $int_type . "_" . $tport . "-Transport_INET-EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Small_Sec_' . $int_type . '_' . $tport . '-Transport_INET-EXT_dtmpl.csv');
			  if ($int_type eq 'static'){
			      $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
			      $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
			  }if ($int_type eq 'dhcp'){
			    prtout( "\n NO AVAILABLE BOOTSTRAP FOR INTERNET TYPE - DHCP. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
			  }
			}if ($tport == 3){
			  # No Available template for 3-transport model... to follow...
			  prtout( "\n NO AVAILABLE SDWAN TEMPLATE and BOOTSTRAP FOR NEW $tport-TRANSPORT DESIGN. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
			}
			}else{
			if ($tport == 2){
			  $outputFile = writeTemplate( $med_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			  $outputFile = writeTemplate( $med_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Sec_' . $tport . '-Transport_dtmpl.csv' );
			  if ($int_type eq 'static'){
			    $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
			    $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			  }if ($int_type eq 'dhcp'){
			    $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			  }
			}else{
			  $outputFile = writeTemplate( $med_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Pri_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Pri_' . $int_type . '_dtmpl.csv');
			  $outputFile = writeTemplate( $med_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Sec_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Sec_dtmpl.csv' );
			  if ($int_type eq 'static'){
			    $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
			    $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			  }if ($int_type eq 'dhcp'){
			    $outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			  }
			}
			}
		}
	if ( $anchor{'site_type'} eq 'L' ) {
		if ($tport == 2){
			$outputFile = writeTemplate( $large_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			$outputFile = writeTemplate( $large_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Sec_' . $tport . '-Transport_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}else{
			$outputFile = writeTemplate( $large_folder . "Trust_RS_Consolidated_" . $rtr_type ."_Large_Pri_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type .  '_Large_Pri_' . $int_type . '_dtmpl.csv');
			$outputFile = writeTemplate( $large_folder . "Trust_RS_Consolidated_" . $rtr_type ."_Large_Sec_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Sec_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}
		}
	if ( $anchor{'site_type'} eq 'RDC' ) {
		if( $anchor{'tloc'} eq 'yes' ){
		if ($tport == 2){
			$outputFile = writeTemplate( $RDC_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Pri_" . $int_type . "_" . $tport . "-Transport_INET-EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Small_Pri_' . $int_type . '_' . $tport . '-Transport_INET-EXT_dtmpl.csv');
			$outputFile = writeTemplate( $RDC_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Sec_" . $int_type . "_" . $tport . "-Transport_INET-EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Small_Sec_' . $int_type . '_' . $tport . '-Transport_INET-EXT_dtmpl.csv');
			if ($int_type eq 'static'){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				prtout( "\n NO AVAILABLE BOOTSTRAP FOR INTERNET TYPE - DHCP. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
			}
		}if ($tport == 3){
			# No Available template for 3-transport model... to follow...
			prtout( "\n NO AVAILABLE SDWAN TEMPLATE and BOOTSTRAP FOR NEW $tport-TRANSPORT DESIGN. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
		}if ($tport == 4){
			# No Available template for 4-transport model... to follow...
			prtout( "\n NO AVAILABLE SDWAN TEMPLATE and BOOTSTRAP FOR NEW $tport-TRANSPORT DESIGN. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
		}
		}else{
		if ($tport == 2){
			$outputFile = writeTemplate( $RDC_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			$outputFile = writeTemplate( $RDC_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Sec_' . $tport . '-Transport_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}else{
			$outputFile = writeTemplate( $RDC_folder . "Trust_RS_Consolidated_" . $rtr_type ."_Large_Pri_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type .  '_Large_Pri_' . $int_type . '_dtmpl.csv');
			$outputFile = writeTemplate( $RDC_folder . "Trust_RS_Consolidated_" . $rtr_type ."_Large_Sec_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Sec_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}
	}
	}
	return $outputFile;
}

}   #This curly braket is the end of the if for legacy site models

####################################################################################
#***********************SUBS FOR LEGACY SITE MODELS END HERE***********************#
####################################################################################

###################################################################################
#***********************SUBS FOR NEW SITE MODELS START HERE***********************#
###################################################################################

else {

our $sitegroup;			# used to target site specific templates
setAnchorValuesNew( $site_code );

# Set other site specific anchor values
# Site branch logic. H,K,G,T, are handled by a generic routine, D,C,N,F need special routines

writeSiteNew() if ( $sitetype =~ /^(D|C|N|F|U)$/ );
writeSiteNew(2) if ( $sitetype eq 'H' );
writeSiteNew(4) if ( $sitetype eq 'K' );
writeSiteNew(8) if ( $sitetype eq 'G' );
writeSiteNew(16) if ( $sitetype eq 'T' );

my $tfinished = getTime();
my @ctrinfo = ($site_code,$tstart,$tfinished, $anchor{'proj_type'},$msid);
my $dbcounter = $dbh->prepare("INSERT INTO nslanwan.counter (SITECODE, DATECOUNTER,FINISHTIME, PRJ_TYPE, USER) values (?,?,?,?,?) ");
$dbcounter->execute(@ctrinfo);
xit(0);

sub setAnchorValuesNew {
	my $site_code = shift;
	my $sitetype = $anchor{'site_type'};

	# Do initial input check
	inputErrorCheckNew();

	# SNMP
	#As part of clean-up, we will be moving this to the main folder under Templates and will stop updating the files under New-standards/Modules
	$anchor{'tg_mls_snmp'}      = smbReadFile("New-Standards/Modules/tg_mls_snmp.txt");
	$anchor{'kh_mls_snmp'}       = smbReadFile("New-Standards/Modules/kh_mls_snmp.txt");

	# Site specific Dir and file values
	my %sitegroups = ( 'H', 'KH', 'K', 'KH', 'G', 'TG', 'T', 'TG' );
	$sitegroup = $sitegroups{$sitetype};
	$sitegroup = 'DCNF' if ( $sitetype !~ /^(T|G|K|H)$/ );

	# FW variables sets
	if ( $anchor{'fw'} eq 'Y' and $sitetype =~ /^(T|G|K|H)$/ ) {
		$anchor{'fw_net'}  = '16/28';
		$anchor{'fw_net'}  = '8/29' if ( $sitetype =~ /^(K|H)$/ );
		$anchor{'fw_gwy'}  = '17';
		$anchor{'fw_gwy'}  = '9' if ( $sitetype =~ /^(K|H)$/ );
		$anchor{'fw_ip'}   = '26';
		$anchor{'fw_ip'}   = '10' if ( $sitetype =~ /^(K|H)$/ );
		$anchor{'fw1_mgmt_ip'} = '72';
		$anchor{'fw1_mgmt_ip'} = '40' if ( $sitetype =~ /^(K|H)$/ );
		$anchor{'fw1_mgmt_ip'} = '136' if ( $sitetype =~ /^T$/ );
		$anchor{'fw2_mgmt_ip'} = '73';
		$anchor{'fw2_mgmt_ip'} = '41' if ( $sitetype =~ /^(K|H)$/ );
		$anchor{'fw2_mgmt_ip'} = '137' if ( $sitetype =~ /^T$/ );
		$anchor{'fw_mls1'} = '18' if ( $sitetype =~ /^(T|G)$/ );
		$anchor{'fw_mls2'} = '19' if ( $sitetype =~ /^(T|G)$/ );
		$anchor{'mls1_fw_config'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_fw_$sitegroup.txt");
		$anchor{'mls2_fw_config'} = smbReadFile("New-Standards/Model-$sitegroup/mls2_fw_$sitegroup.txt")
		if ( $sitetype =~ /^(T|G)$/ );
	} else {
		$anchor{'mls1_fw_int'}    = " description reserved\r\n shutdown";
		$anchor{'mls1_fw_vl'}     = '!';
		$anchor{'mls1_fw_config'} = '!';
		$anchor{'mls2_fw_int'}    = " description reserved\r\n shutdown";
		$anchor{'mls2_fw_vl'}     = '!';
		$anchor{'mls2_fw_config'} = '!';
	}

	# Core Device names
	my $mdf_flr =
	$anchor{'mdf_bldg'} . sprintf( "%02d", $anchor{'mdf_flrnumber'} );
	$anchor{'cis1_name'} = 'cis' . $site_code . $mdf_flr . 'a01';    # site types: all
	# C and F sites
	$anchor{'cis2_name'} = '***DefnErr***';
	# site types: all but C and F
	$anchor{'cis2_name'} = 'cis' . $site_code . $mdf_flr . 'a02'
	if ( $sitetype !~ /^(C|F|U)$/ );
	if ( $sitetype =~ /^(T|G|K|H)$/ ) {                              # site types: all but D, C, N, and F
		$anchor{'mls1_name'} = 'mls' . $site_code . $mdf_flr . 'a01';
		# K and H sites
		$anchor{'mls2_name'} = '***DefnErr***';
		# T and G sites
		$anchor{'mls2_name'} = 'mls' . $site_code . $mdf_flr . 'a02'
		if ( $sitetype =~ /^(T|G)$/ );
	} else {    # D, C, N, F sites may need this set for the vgc config file
		$anchor{'mls1_name'} = ( split( /,/, $StackList[0] ) )[0];
		$anchor{'mls2_name'} = ( split( /,/, $StackList[0] ) )[0];
	}
	$anchor{'cgw1_name'} = 'cgw' . $site_code . $mdf_flr . 'a01';    # site types: U
	$anchor{'cgw2_name'} = 'cgw' . $site_code . $mdf_flr . 'a02';    # site types: U

	$anchor{'netscout1_name'} = 'mon' . $site_code . $mdf_flr . 'a01'; #Netscout name
	$anchor{'sbc1_oracle_name'} = 'acm' . $site_code . 'mssbc' . '01a'; #Oracle SBC name
	$anchor{'sbc2_oracle_name'} = 'acm' . $site_code . 'mssbc' . '01b'; #Oracle SBC name


	# Site type specific subnets
	my $oa = $anchor{'loop_oct1'};
	my $ob = $anchor{'loop_oct2'};
	my $oc = $anchor{'loop_oct3'};
	$anchor{'loop_subnet'} = "$oa.$ob.$oc";
	$anchor{'data_subnet_1'} = "$oa.$ob.$oc";

	#Declaring variable for U site
	my $infrasubnet = $anchor{'Infrasubnet'};
	my ($o1, $o2, $o3, $o4) = split(/\./, $infrasubnet);
	my $ucis = $o4 + 1;
	$anchor{'ucis'} = "$oa.$ob.$oc.$ucis";
	my $ucgw1 = $o4 + 2;
	$anchor{'ucgw1'} = "$oa.$ob.$oc.$ucgw1";
	my $ucgw2 = $o4 + 3;
	$anchor{'ucgw2'} = "$oa.$ob.$oc.$ucgw2";
	my $utgw = $o4 + 4;
	$anchor{'utgw'} = "$oa.$ob.$oc.$utgw";
	my $utad = $o4 + 5;
	$anchor{'utad'} = "$oa.$ob.$oc.$utad";
	my $uena = $o4 + 8;
	$anchor{'uena'} = "$oa.$ob.$oc.$uena";
	my $uegw = $o4 + 9;
	$anchor{'uegw'} = "$oa.$ob.$oc.$uegw";
	my $udna = $o4 + 16;
	$anchor{'udna'} = "$oa.$ob.$oc.$udna";
	my $udgw = $o4 + 17;
	$anchor{'udgw'} = "$oa.$ob.$oc.$udgw";

	my %sitemask = ( 'T', '255.255.224.0', 'G', '255.255.240.0', 'K', '255.255.248.0', 'H', '255.255.252.0',
					'D', '255.255.254.0', 'C', '255.255.254.0', 'N', '255.255.255.0', 'F', '255.255.255.0', 'U', '255.255.255.224' );
	my %sitecidr = ( 'T', '19', 'G', '20', 'K', '21', 'H', '22', 'D', '23', 'C', '23', 'N', '24', 'F', '24', 'U', '27' );
	my %siteName = ( 'RDC', 'RDC', 'T', 'Tera', 'G', 'Giga', 'K', 'Kilo', 'H', 'Hecto',
					'D', 'Deci', 'C', 'Centi', 'N', 'Nano', 'F', 'Femto', 'U', 'Micro' );

	prtout("SITE TYPE: $sitetype");

	$anchor{'site_mask'} = $sitemask{$sitetype};
	$anchor{'site_cidr'} = $sitecidr{$sitetype};
	$anchor{'site_name'} = $siteName{$sitetype};

	$anchor{'mdf_flr'} = substr( $anchor{'cis1_name'}, 8, 3 );

	# TAD
	$anchor{'tad1_name'}   = 'tad' . $site_code . $mdf_flr . 'a01';    # site types: all
	$anchor{'tad1_subnet'} = "$oa.$ob.$oc.16/29";
	$anchor{'tad1_ip'} = "$oa.$ob.$oc.20";
	$anchor{'tad1_mask'} = "255.255.255.248";
	$anchor{'tad1_gwy'} = '19';
	$anchor{'tad1_subnet'} = "$oa.$ob.$oc.56/30" if ( $sitetype =~ /^(T|G)$/ );
	$anchor{'tad1_ip'} = "$oa.$ob.$oc.57" if ( $sitetype =~ /^(T|G)$/ );
	$anchor{'tad1_mask'} = "255.255.255.252" if ( $sitetype =~ /^(T|G)$/ );
	$anchor{'tad1_gwy'} = '58' if ( $sitetype =~ /^(T|G)$/ );

	# Server Subnet
	if ( $sitetype =~ /^(K|H)$/ ) {
		$anchor{'svr_gwy'} = '33';
		$anchor{'svr_mask'} = '255.255.255.224';
		$anchor{'svr_cidr'} = '27';
	} elsif ( $sitetype eq 'G' ) {
		$anchor{'svr_gwy'} = '65';
		$anchor{'svr_mls1'} = '66';
		$anchor{'svr_mls2'} = '67';
		$anchor{'svr_mask'} = '255.255.255.192';
		$anchor{'svr_cidr'} = '26';
	} elsif ( $sitetype eq 'T' ) {
		$anchor{'svr_gwy'} = '129';
		$anchor{'svr_mls1'} = '130';
		$anchor{'svr_mls2'} = '131';
		$anchor{'svr_mask'} = '255.255.255.128';
		$anchor{'svr_cidr'} = '25';
	}

	# Data subnets for D, C, N and F sites
	if ( $sitetype =~ /^(D|C)$/ ) {
		$oc++; $anchor{'data_subnet_1'} = "$oa.$ob.$oc";
		$anchor{'data_stk_oct'} = '1';
		$anchor{'data_stk_mask'} = '255.255.255.0';
		$anchor{'data_stk_cidr'} = '24';
		# Convert data subnet to hex for Aruba vlan201 secondary IP
		$anchor{'data_stk_sec'} = '4';
		my ($h1,$h2,$h3,$h4) = (sprintf("%02X", $oa), sprintf("%02X", $ob), sprintf("%02X", $oc), sprintf("%02X", $anchor{'data_stk_sec'}));
		$anchor{'oct_to_hex'} = ("$h1:$h2:$h3:$h4");
		prtout("$anchor{'oct_to_hex'}");
	} elsif ( $sitetype =~ /^(N|F)$/ ) {
		$anchor{'data_stk_oct'} = '129';
		$anchor{'data_stk_mask'} = '255.255.255.128';
		$anchor{'data_stk_cidr'} = '25';
		# Convert data subnet to hex for Aruba vlan201 secondary IP
		$anchor{'data_stk_sec'} = '132';
		my ($h1,$h2,$h3,$h4) = (sprintf("%02X", $oa), sprintf("%02X", $ob), sprintf("%02X", $oc), sprintf("%02X", $anchor{'data_stk_sec'}));
		$anchor{'oct_to_hex'} = ("$h1:$h2:$h3:$h4");
		prtout("$anchor{'oct_to_hex'}");
	}

	$anchor{'eac_subnet_wireless'} = '***DefnErr***';

	# Router types
	# - If default router selected, determine router based on subrate
	# - If specific router is selected, use that setting
	if ( $anchor{'router_seltype'} =~ /^(4321|4331|4351|4451|4461|ASR|C8200-1N-4T|C8300-1N1S|C8300-2N2S|C1111X-8P|C1161X-8P|C1121X-8P|C8500-12X4QC)$/ ) {
		$anchor{'router_type'} = $anchor{'router_seltype'};
	} elsif ( $sitetype =~ /^(T|G|K|H)$/ ) {    #TGKH Default router code due to 4321 lack of ports.
		if ( $anchor{'pri_circuit_type'} =~ /^(T1|2xMLPPP|3xMLPPP|4xMLPPP|2xMLPPP-E1|3xMLPPP-E1|4xMLPPP-E1)$/ ) {
			$anchor{'router_type'} = '4351';
		} elsif ( $anchor{'pri_circuit_type'} =~ /^(Metro|MPLS)_Ethernet$/ ) {
			my $subrate = $anchor{'subrate'};
			if ( $subrate >= 50000 and $subrate < 200000 ) {                  #
				$anchor{'router_type'} = '4351';
			} elsif ( $subrate >= 200000 and $subrate <= 1000000 ) {
				$anchor{'router_type'} = '4451';
			} else {
				$anchor{'router_type'} = '4351';
			}
		} else {
			$anchor{'router_type'} = '4451';
		}
	} else {    # Set to default values
		if ( $anchor{'pri_circuit_type'} =~ /^(T1|2xMLPPP|3xMLPPP|4xMLPPP|2xMLPPP-E1|3xMLPPP-E1|4xMLPPP-E1)$/ ) {

			$anchor{'router_type'} = '4321';    # future standard - noted on 2017-04-20
		} elsif ( $anchor{'pri_circuit_type'} =~ /^(Metro|MPLS)_Ethernet$/ ) {
			my $subrate = $anchor{'subrate'};
			if ( $subrate >= 50000 and $subrate < 200000 ) {    #
				$anchor{'router_type'} = '4351';
			} elsif ( $subrate >= 200000 and $subrate <= 1000000 ) {
				$anchor{'router_type'} = '4451';
			} else {

				$anchor{'router_type'} = '4321';                # future standard - noted on 2017-04-20
			}
		} else {
			$anchor{'router_type'} = '4451';                    # was 2951 - 2017-04-20
		}
	}
	# TLOC Interfaces for T and G
	$anchor{'tloc_upl1'} = 'Gi0/1/0.40';
	$anchor{'tloc_upl1'} = 'TE0/0/4.40' if ( $anchor{'router_type'} eq '4461');
	$anchor{'tloc_upl2'} = 'Gi0/0/2.40';
	$anchor{'tloc_upl2'} = 'TE0/0/4.40' if ( $anchor{'router_type'} eq '4461');

	# Site wildcard mask
	my %wcmask = (
				   'U', '0.0.0.31', 'F', '0.0.0.255', 'N', '0.0.0.255', 'C', '0.0.1.255', 'D', '0.0.1.255',
				   'H', '0.0.3.255', 'K', '0.0.7.255', 'G', '0.0.15.255', 'T', '0.0.31.255'
	);
	if ( defined $wcmask{$sitetype} ) {
		$anchor{'wildcard_mask'} = $wcmask{$sitetype};
	} else {
		prtout( "Error: Selected site type '$sitetype' not currently supported in this version of Configurator.",
				"Please contact the developers concerning the availability of running Configurator with this model." );
		xit(1);
	}

	# Visio template updates
	getVisioTemplatesNew($sitetype);

	# New site model K|H|T|G spare amt
	if ( $sitetype =~ /^(K|H)$/ ){
 	$anchor{'spare_amt'}   = 1;
 	}elsif ( $sitetype =~ /^(T|G)$/ ){
 	$anchor{'spare_amt'}   = 2;
	}

	#SBC configuration
	if ($anchor{'sbc'} eq 'yes'){
	prtout("Updating SBC Configuration");
	if ( $sitetype =~ /^(G|T)$/ ) {
		$anchor{'sbc_mls_int1'} = 'GigabitEthernet5/0/15';
		$anchor{'sbc_mls_int2'} = 'GigabitEthernet5/0/16';
		$anchor{'sbc_mls_int3'} = 'GigabitEthernet5/0/17';
		# $anchor{'sbc_mls_int1'} = 'GigabitEthernet2/0/15';
		# $anchor{'sbc_mls_int2'} = 'GigabitEthernet2/0/16';
		# $anchor{'sbc_mls_int3'} = 'GigabitEthernet2/0/17';
	} elsif ( $sitetype =~ /^(H|K)$/ ) {
		$anchor{'sbc_mls_int1'} = 'GigabitEthernet1/0/15';
		$anchor{'sbc_mls_int2'} = 'GigabitEthernet2/0/15';
		$anchor{'sbc_mls_int3'} = 'GigabitEthernet1/0/16';
		$anchor{'sbc_mls_int4'} = 'GigabitEthernet2/0/16';
		$anchor{'sbc_mls_int5'} = 'GigabitEthernet1/0/17';
		$anchor{'sbc_mls_int6'} = 'GigabitEthernet2/0/17';
	}

	# Calculating sbc subnets
	# my %sbc397gw  = ( 'T', '129',       'G', '129',       'K', '129',       'H', '129'     );
	# my %sbc398gw  = ( 'T', '225',       'G', '225',       'K', '225',       'H', '113'     );

	my ( $oa, $ob, $oc ) = split( /\./, $anchor{'loop_subnet'} );
	my $sbco;
	$sbco = $oc if ( $sitetype eq 'H' );
	$sbco = ($oc + 6) if ( $sitetype eq 'K' );
	$sbco = ($oc + 11) if ( $sitetype eq 'G' );
	$sbco = ($oc + 1) if ( $sitetype eq 'T' );
	$anchor{'sbc_sub'} = "$oa.$ob.$sbco";
	$anchor{'sbc_397gw'} = "$oa.$ob.$sbco.129";
	$anchor{'sbc_398gw'} = "$oa.$ob.$sbco.225";
	$anchor{'sbc_398gw'} = "$oa.$ob.$sbco.113" if ( $sitetype eq 'H' );
	$anchor{'sbc_397_ip1'} = "$oa.$ob.$sbco.130";
	$anchor{'sbc_397_ip2'} = "$oa.$ob.$sbco.131";
	$anchor{'sbc_398_ip1'} = "$oa.$ob.$sbco.226";
	$anchor{'sbc_398_ip2'} = "$oa.$ob.$sbco.227";
	$anchor{'sbc_398_lastip'} = "$oa.$ob.$sbco.238";
	$anchor{'sbc_398_lastip'} = "$oa.$ob.$sbco.126" if ( $sitetype eq 'H' );
	$anchor{'sbc397_mask'} = "255.255.255.192";
	$anchor{'sbc398_mask'} = "255.255.255.240";

	# my $anchor{'mdf_flr'} =
	#   $anchor{'mdf_bldg'} . sprintf( "%02d", $anchor{'mdf_flrnumber'} );

	if ($sitetype =~ /^(T|G)$/){
	$anchor{'mls1_sbc_tg'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_sbc_TG.txt");
	$anchor{'mls2_sbc_tg'} = smbReadFile("New-Standards/Model-$sitegroup/mls2_sbc_TG.txt");
	#Oracle SBC TAD port config
	$anchor{'set_deviceport_14'} = smbReadFile("New-Standards/Misc/oracle_sbc1.txt");
	$anchor{'set_deviceport_15'} = smbReadFile("New-Standards/Misc/oracle_sbc2.txt");
	}if($sitetype =~ /^(K|H)$/){
	$anchor{'mls1_sbc_kh'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_sbc_KH.txt");
	#Oracle SBC TAD port config
	$anchor{'set_deviceport_14'} = smbReadFile("New-Standards/Misc/oracle_sbc1.txt");
	$anchor{'set_deviceport_15'} = smbReadFile("New-Standards/Misc/tadport_15.txt");
	}
	}
	elsif($anchor{'sbc'} eq 'no'){
	if ($sitetype =~ /^(T|G)$/){
	$anchor{'mls1_sbc_tg'} = '!';
	$anchor{'mls2_sbc_tg'} = '!';
	#Oracle SBC TAD port config
	$anchor{'set_deviceport_14'} = smbReadFile("New-Standards/Misc/tadport_14.txt");
	$anchor{'set_deviceport_15'} = smbReadFile("New-Standards/Misc/tadport_15.txt");
	}if($sitetype =~ /^(K|H)$/){
	$anchor{'mls1_sbc_kh'} = '!';
	#Oracle SBC TAD port config
	$anchor{'set_deviceport_14'} = smbReadFile("New-Standards/Misc/tadport_14.txt");
	$anchor{'set_deviceport_15'} = smbReadFile("New-Standards/Misc/tadport_15.txt");
	}
	}
	
	##Additional anchor values for SDWAN Next-Gen Templates
	#Below anchor values are for Thousand Eyes, TLOC Neighbor IP, and Dummy placeholder for TLOC interfaces
	#***NOTE: 'tloc' is for the subnet of the mpls ckt located on R1
	#***NOTE: 'tloc2' is for the subnet of the mpls ckt located on R2
	my $loop_subnet = $anchor{'loop_subnet'};
	if ( $anchor{'site_type'} =~ /^(D|C|N|F)$/ ) {
		$anchor{'theyes_mgmt'} = "$loop_subnet.6/30";
		$anchor{'theyes_gwip'} = "$loop_subnet.5";
		$anchor{'biz_tloc_ip1'} = "$loop_subnet.25";
		$anchor{'biz_tloc_ip2'} = "$loop_subnet.26";
		$anchor{'prvt1_tloc_na'} = "$loop_subnet.12/30";
		$anchor{'prvt1_tloc_ip1'} = "$loop_subnet.13";
		$anchor{'prvt1_tloc_ip2'} = "$loop_subnet.14";
		$anchor{'pub_prvt2_bgp_prfx'} = "$loop_subnet.28/30";
		$anchor{'pub_prvt2_tloc_ip1'} = "$loop_subnet.29";
		$anchor{'pub_prvt2_tloc_ip2'} = "$loop_subnet.30";
	}if ( $anchor{'site_type'} =~ /^(K|H)$/ ) {
		$anchor{'theyes_mgmt'} = "$loop_subnet.14/30";
		$anchor{'theyes_gwip'} = "$loop_subnet.13";
		$anchor{'biz_tloc_ip1'} = "$loop_subnet.25";
		$anchor{'biz_tloc_ip2'} = "$loop_subnet.26";
		$anchor{'prvt1_tloc_na'} = "$loop_subnet.96/30";
		$anchor{'prvt1_tloc_ip1'} = "$loop_subnet.97";
		$anchor{'prvt1_tloc_ip2'} = "$loop_subnet.98";
		$anchor{'pub_prvt2_bgp_prfx'} = "$loop_subnet.28/30";
		$anchor{'pub_prvt2_tloc_ip1'} = "$loop_subnet.29";
		$anchor{'pub_prvt2_tloc_ip2'} = "$loop_subnet.30";
	}if ( $anchor{'site_type'} =~ /^(G)$/ ) {
		$anchor{'theyes_mgmt'} = "$loop_subnet.202/30";
		$anchor{'theyes_gwip'} = "$loop_subnet.201";
		$anchor{'biz_tloc_ip1'} = "$loop_subnet.53";
		$anchor{'biz_tloc_ip2'} = "$loop_subnet.54";
		$anchor{'prvt1_tloc_na'} = "$loop_subnet.192/30";
		$anchor{'prvt1_tloc_ip1'} = "$loop_subnet.193";
		$anchor{'prvt1_tloc_ip2'} = "$loop_subnet.194";
		$anchor{'pub_prvt2_bgp_prfx'} = "$loop_subnet.60/30";
		$anchor{'pub_prvt2_tloc_ip1'} = "$loop_subnet.61";
		$anchor{'pub_prvt2_tloc_ip2'} = "$loop_subnet.62";
	}if ( $anchor{'site_type'} =~ /^(T)$/ ) {
		$anchor{'theyes_mgmt'} = "$loop_subnet.74/30";
		$anchor{'theyes_gwip'} = "$loop_subnet.73";
		$anchor{'biz_tloc_ip1'} = "$loop_subnet.53";
		$anchor{'biz_tloc_ip2'} = "$loop_subnet.54";
		$anchor{'prvt1_tloc_na'} = "$loop_subnet.64/30";
		$anchor{'prvt1_tloc_ip1'} = "$loop_subnet.65";
		$anchor{'prvt1_tloc_ip2'} = "$loop_subnet.66";
		$anchor{'pub_prvt2_bgp_prfx'} = "$loop_subnet.60/30";
		$anchor{'pub_prvt2_tloc_ip1'} = "$loop_subnet.61";
		$anchor{'pub_prvt2_tloc_ip2'} = "$loop_subnet.62";
	}
	#prtout("TH Eyes Mgmt: $anchor{'theyes_mgmt'}");
	#prtout("TH Eyes GW: $anchor{'theyes_gwip'}");

	if( $anchor{'pub_provider'} =~ m/Internet/i ){
		$anchor{'pub_tloc_stat'} = 'FALSE';
		$anchor{'prvt2_tloc2_stat'} = 'TRUE';
		$anchor{'pub_tloc_ip1'} = $anchor{'pub_prvt2_tloc_ip1'};
		$anchor{'pub_tloc_ip2'} = $anchor{'pub_prvt2_tloc_ip2'};
		$anchor{'r1_pub_tloc_desc'} = "TLOC Ext. PUBLIC-INTERNET Interface from $anchor{'cis2_name'}";
		$anchor{'r1_prvt2_tloc2_desc'} = 'UNUSED';
		$anchor{'prvt2_tloc_ip1'} = '169.254.0.1';
		if($anchor{'transport'} == 2){
			$anchor{'r2_color_tloc'} 	= 'mpls';
			$anchor{'r2_carrier_tloc'}	= 'carrier1';
		}if($anchor{'transport'} == 3){
			if(($anchor{'r1_provider'} ne 'ATT')){
				$anchor{'r2_color_tloc'} 	= 'mpls';
				$anchor{'r2_carrier_tloc'}	= 'carrier1';
			}if(($anchor{'r1_provider'} ne 'VZ')){
				$anchor{'r2_color_tloc'} 	= 'private1';
				$anchor{'r2_carrier_tloc'}	= 'carrier2';
			}if(($anchor{'r1_provider'} ne 'LUMEN')){
				$anchor{'r2_color_tloc'} 	= 'private2';
				$anchor{'r2_carrier_tloc'}	= 'carrier5';
			}
		}
	}if( $anchor{'r2_provider'} =~ /^(ATT|VZ|LUMEN)$/ ){
		$anchor{'pub_tloc_stat'} = 'TRUE';
		$anchor{'prvt2_tloc2_stat'} = 'FALSE';
		$anchor{'prvt2_tloc_ip1'} = $anchor{'pub_prvt2_tloc_ip1'};
		$anchor{'prvt2_tloc_ip2'} = $anchor{'pub_prvt2_tloc_ip2'};
		$anchor{'r1_prvt2_tloc2_desc'} = "TLOC Ext. $anchor{'r2_color_uc'} Interface from $anchor{'cis2_name'}";
		$anchor{'r1_pub_tloc_desc'} = 'UNUSED';
		$anchor{'pub_tloc_ip1'} = '169.254.0.1';
		$anchor{'r2_color_tloc'} = $anchor{'r2_color'};
		$anchor{'r2_carrier_tloc'} = $anchor{'r2_carrier'};
	}
}

sub getVisioTemplatesNew {
return unless ($anchor{'proj_type'} eq 'build');
	my $sitetype = shift;
	prtout("Downloading Visio template for site type '$sitetype'\n");
	if ( $sitetype =~ /^(T|G|K|H)$/ ) {
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/generic-tgkh/master-tgkh-template.vdx",       "$ROOTDIR/master-tgkh-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/generic-tgkh/master-tgkh-template-clean.vdx", "$ROOTDIR/master-tgkh-template-clean.vdx" );
	} elsif ( $sitetype =~ /^(D|N)$/ ) {
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/dn_site/master-dn-template.vdx",       "$ROOTDIR/master-dn-template.vdx" );
		#smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/dn_site/master-dn-template.vsdx",       "$ROOTDIR/master-dn-template.vsdx" );
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/dn_site/master-dn-template-clean.vdx", "$ROOTDIR/master-dn-template-clean.vdx" );
	} elsif ( $sitetype =~ /^(C|F)$/ ) {
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/cf_site/master-cf-template.vdx",       "$ROOTDIR/master-cf-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/cf_site/master-cf-template-clean.vdx", "$ROOTDIR/master-cf-template-clean.vdx" );
	} elsif ( $sitetype =~ /^(U)$/ ) {
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/u_site/master-u-template.vdx",       "$ROOTDIR/master-u-template.vdx" );
		smbGet( "$SMB_TEMPLATE_DIR/New-Standards/Visio/u_site/master-u-template-clean.vdx", "$ROOTDIR/master-u-template-clean.vdx" );
	} else {
		prtout( "Unknown site type! Please check the site type and resubmit, or contact the developer ",
				" if the site type is correct." );
		exit;
	}
}

sub inputErrorCheckNew {
	my $numstack = scalar(@StackList);
	my $sitetype = $anchor{'site_type'};
	my $firewall = $anchor{'fw'};
	my $xsubnet  = $anchor{'xsubnet'};
	my $routertype = $anchor{'router_seltype'};

	if ( $sitetype !~ /^(T|G|K|H|D|C|N|F|U)$/ ){
		prtout("Site Model type '$sitetype' is not valid");
		xit(1);
	}
	if ( $sitetype =~ /^(D|C|N|F|U)$/ and $firewall eq 'Y' ) {
		prtout(
				"Currently, Configurator does not support $sitetype sites with a firewall.",
				"You will need to manually configure the firewall aspect of the $sitetype site.",
				"Please discuss with Steve or Dan if you require a further explanation",
				"Site Type: $sitetype",
				"Firewall Check: $firewall"
		);
		xit(1);
	}
	if ( $routertype =~ /^(3945E|3945||2951|)$/ ) {
		prtout(
				"Currently, Configurator does not support $sitetype sites with a $routertype Router selected.",
				"Please discuss with Steve or Dan if you require a further explanation.",
				"Site Type: $sitetype",
				"Router Check: $routertype"
		);
		xit(1);
	}
}

sub wirelessAPNew {
	return unless ( $anchor{'wlan'} eq 'Y' );
	my $sitetype = $anchor{'site_type'};
	my $readFile;
	$anchor{'flex_controllers'} = '!';
	if ( $sitetype =~ /^(D|C|N|F|U)$/ ) {
		# Flex controllers at DCs
		$anchor{'wlc1_mgmt_ip'} = '<FILL_IN_' . $anchor{'wireless_region'} . '_IP>';
		$anchor{'wlc1_name'} = '<FILL_IN_' . $anchor{'wireless_region'} . '_NAME>';
		if( $anchor{'wireless_region'} eq 'WEST' ) { # Sets west flex controllers
			$anchor{'wlc_ter_ip'} = '10.141.58.144';
			$anchor{'wlc_ter_name'}    = 'wlcMN053bkpa21';
			$anchor{'flex_controllers'} = smbReadFile("New-Standards/Wireless/AP_flex_west.txt");
		}
		else { # Sets east and international flex controllers
			$anchor{'wlc_ter_ip'} = '10.141.62.149';
			$anchor{'wlc_ter_name'}    = 'wlcMN011bkpa21';
			$anchor{'flex_controllers'} = smbReadFile("New-Standards/Wireless/AP_flex_east.txt");
		}
	}
	# Data wireless subnets for D and C sites
	if ( $sitetype =~ /^(D|C)$/ ) {
		$anchor{'wlan_subnet_i'} = "$anchor{'loop_subnet'}.128";
		$anchor{'wlan_gwyip_i'} = "$anchor{'loop_subnet'}.129";
		prtout("VLAN110: $anchor{'wlan_gwyip_i'}");
		$anchor{'wlan_subnet_i_mask'} = '255.255.255.128';
		$anchor{'wlan_subnet_i_mask_nexus'} = '/25';
		$readFile = "New-Standards/Model-$sitegroup/stk_stl_wlan.txt";
		$readFile = "New-Standards/Model-$sitegroup/aruba/stk_stl_wlan.txt"
		if ($anchor{'stack_vendor'} eq 'aruba');
		$anchor{'dc_wlan'} = smbReadFile($readFile);
	} else { ( $anchor{'dc_wlan'} = '!' ) };
	prtout("Updating Wireless AP list file");
	writeTemplate( "New-Standards/Wireless/AP_data.txt", $anchor{'site_code'} . '-WAPs.txt' );

	#AP placement Attestation
	writeTemplate( "New-Standards/Wireless/Attestation/AP Placement Attestation.xml", $anchor{'site_code'} . ' - AP Placement Attestation.doc' );

}

sub wirelessControllerNew {
	return unless ( $anchor{'wlan'} eq 'Y' );
	my $sitetype = $anchor{'site_type'};
	if ( $sitetype =~ /^(G|T)$/ ) {
		$anchor{'wireless_lag'} = 'lag disable';
	} elsif ( $sitetype =~ /^(H|K)$/ ) {
		$anchor{'wireless_lag'} = 'lag enable';
	}
	prtout("Updating Wireless Controller Configuration");
	if ( $sitetype =~ /^(G|T)$/ ) {
		$anchor{'wlc1_mls_int'}  = 'GigabitEthernet5/0/9';
		$anchor{'wlc1_mls_int1'} = 'TenGigabitEthernet1/0/23';
		$anchor{'wlc1_mls_int2'} = 'TenGigabitEthernet1/0/24';
		$anchor{'wlc2_mls_int1'} = 'TenGigabitEthernet1/0/23';
		$anchor{'wlc2_mls_int2'} = 'TenGigabitEthernet1/0/24';
	} elsif ( $sitetype =~ /^(H|K)$/ ) {
		$anchor{'wlc_mls_int1'}  = 'GigabitEthernet1/0/9';
		$anchor{'wlc_mls_int2'}  = 'GigabitEthernet2/0/9';
		$anchor{'wlc1_mls_int1'} = 'TenGigabitEthernet1/1/8';
		$anchor{'wlc1_mls_int2'} = 'TenGigabitEthernet2/1/8';
		$anchor{'wlc2_mls_int1'} = 'TenGigabitEthernet1/1/7';
		$anchor{'wlc2_mls_int2'} = 'TenGigabitEthernet2/1/7';
	}

	# Calculating wireless subnets
	my %wgi  = ( 'T', '1',       'G', '1',       'K', '1',       'H', '129'     );
	my %wwi  = ( 'T', '4',       'G', '4',       'K', '4',       'H', '132'     );
	my %wse  = ( 'T', '0',       'G', '128',     'K', '64',      'H', '64' 	    );
	my %wge  = ( 'T', '1',       'G', '129',     'K', '65',      'H', '65'      );
	my %wm1e = ( 'T', '2',       'G', '130',     'K', '',        'H', '' 	    );
	my %wm2e = ( 'T', '3',       'G', '131',     'K', '',        'H', '' 	    );
	my %wwe  = ( 'T', '4',       'G', '132',     'K', '68',      'H', '68' 	    );
	my %wsim = ( 'T', '252.0',   'G', '254.0',   'K', '255.0',   'H', '255.128' );
	my %wsin = ( 'T', '22',      'G', '23',      'K', '24',      'H', '25'      );
	my %wsem = ( 'T', '255.128', 'G', '255.192', 'K', '255.224', 'H', '255.224' );
	my %wsen = ( 'T', '25',      'G', '26',      'K', '27',      'H', '27'      );
	my %wmi  = ( 'T', '134',     'G', '70',      'K', '38',      'H', '38'      );
	my %wmi2  = ( 'T', '135',     'G', '71',      'K', '39',      'H', '39'      );
	my ( $oa, $ob, $oc ) = split( /\./, $anchor{'loop_subnet'} );
   	my ( $doc, $eoc );
	$doc = ( $oc + 32 ) - 4 if ( $sitetype eq 'T' );
	$doc = ( $oc + 16 ) - 2 if ( $sitetype eq 'G' );
	$doc = $oc + 7 if ( $sitetype eq 'K' );
	$doc = $oc + 1 if ( $sitetype eq 'H' );
	$eoc = $oc;
	$eoc = $oc + 1 if ( $sitetype eq 'T' );
	$anchor{'eac_subnet_wireless'} 	    = '***DefnErr***';
	$anchor{'wlan_subnet_i'} 	    	= "$oa.$ob.$doc";
	$anchor{'wlan_gwyip_i'} 	    	= "$oa.$ob.$doc.$wgi{$sitetype}";
	$anchor{'wlan_wlcip_i'} 	    	= "$oa.$ob.$doc.$wwi{$sitetype}";
	$anchor{'wlan_subnet_e'} 	    	= "$oa.$ob.$eoc.$wse{$sitetype}";
	$anchor{'wlan_gwyip_e'} 	    	= "$oa.$ob.$eoc.$wge{$sitetype}";
	$anchor{'wlan_mls1ip_e'} 	    	= "$oa.$ob.$eoc.$wm1e{$sitetype}";
	$anchor{'wlan_mls2ip_e'}	    	= "$oa.$ob.$eoc.$wm2e{$sitetype}";
	$anchor{'wlan_wlcip_e'} 	    	= "$oa.$ob.$eoc.$wwe{$sitetype}";
	$anchor{'wlan_subnet_i_mask'}       = "255.255.$wsim{$sitetype}";
	$anchor{'wlan_subnet_i_mask_nexus'} = $wsin{$sitetype};
	$anchor{'wlan_subnet_e_mask'}       = "255.255.$wsem{$sitetype}";
	$anchor{'wlan_subnet_e_mask_nexus'} = $wsen{$sitetype};
	$anchor{'wlc1_mgmt_ip'} 			= "$oa.$ob.$oc.$wmi{$sitetype}";
	$anchor{'wlc2_mgmt_ip'} 			= "$oa.$ob.$oc.$wmi2{$sitetype}";

	my $mdf_flr =
	  $anchor{'mdf_bldg'} . sprintf( "%02d", $anchor{'mdf_flrnumber'} );
	$anchor{'wlc1_name'} = 'wlc' . $anchor{'site_code'} . $mdf_flr . 'a01';
	$anchor{'wlc2_name'} = 'wlc' . $anchor{'site_code'} . $mdf_flr . 'a02';

	if( $anchor{'wireless_region'} eq 'EAST' ){ # Sets US east tertiary controller
	  $anchor{'wlc_ter_ip'} = '10.141.58.144';
    $anchor{'wlc_ter_name'}    = 'wlcMN053bkpa21';
	}
	elsif( $anchor{'wireless_region'} eq 'WEST' ){ # Sets US west tertiary controller
		$anchor{'wlc_ter_ip'} = '10.141.62.149';
		$anchor{'wlc_ter_name'}    = 'wlcMN011bkpa21';
	}
	else {
		$anchor{'wlc_ter_ip'} = '10.177.22.9'; # Sets tertiary for all non-US
		$anchor{'wlc_ter_name'}    = 'wlcMN053bkpa01';
	}

	# Region data
	my $state = '';
	if ( $anchor{'region'} eq 'USA' ) {
		prtout("Region: USA");
		$state = $anchor{'state'};
		$state =~ tr/a-z/A-Z/;
		if ( !( defined $hostnetflow{$state} ) or $hostnetflow{$state} eq '' ) {
			prtout(
					"There appears to be a problem in looking up host information for the state selected below.",
					"State: $state",
					"Please verify the state abbreviation is correct on the NSLANWAN website and rerun Configurator.",
					"If there continues to be issues, please contact Steve or Dan."
			);
			xit(1);
		}
	} else {
		prtout("Region: non-USA: $anchor{'region'}");
		$state = $anchor{'region'};
	}

	# Temporary solution to WLC snmp limitation, to be updated once all old snmp are removed
	my @east = ('AL','AR','CT','DE','FL','GA','IA','IL','IN','KY','LA','MA','MD','ME','MI','MO','MS','NC','NH','NJ','NY','OH',
	'PA','RI','SC', 'TN', 'VA', 'VT', 'WA', 'WI', 'WV', 'Asia (India and West)','Pacific (East of India)','Europe','Canada-CST');
	my @west = ('AK','AZ','CA','CO','HI','ID','KS','MN','MT','ND','NE','NM','NV','OK','OR','SD','TX', 'UT','WA','WY');

	if($state ~~ @east){
		$anchor{'snmp_host1'} = "10.208.154.191";
		$anchor{'snmp_host2'} = "10.177.72.106";
		$anchor{'snmp_host3'} = "10.122.72.169";
		$anchor{'snmp_host4'} = "10.87.57.127";
		$anchor{'snmp_host5'} = "10.86.186.96";
		$anchor{'snmp_host6'} = "10.86.142.225";
		$anchor{'wlc_snmp_name1'} = "cpieast.uhc.com";
		$anchor{'wlc_snmp_name2'} = "apslp0722";
		$anchor{'wlc_snmp_name3'} = "apsls0208";
		$anchor{'wlc_snmp_name4'} = "rp000057185";
		$anchor{'wlc_snmp_name5'} = "rn000057183";
		$anchor{'wlc_snmp_name6'} = "rp000073778";
		}

	if($state ~~ @west){
		$anchor{'snmp_host1'} = "10.208.155.88";
		$anchor{'snmp_host2'} = "10.177.72.148";
		$anchor{'snmp_host3'} = "10.122.72.188";
		$anchor{'snmp_host4'} = "10.87.57.127";
		$anchor{'snmp_host5'} = "10.86.186.96";
		$anchor{'snmp_host6'} = "10.29.74.248";
		$anchor{'wlc_snmp_name1'} = "cpiwest.uhc.com";
		$anchor{'wlc_snmp_name2'} = "apslp0724";
		$anchor{'wlc_snmp_name3'} = "apsls0210";
		$anchor{'wlc_snmp_name4'} = "rp000057185";
		$anchor{'wlc_snmp_name5'} = "rn000057183";
		$anchor{'wlc_snmp_name6'} = "vp000054652";
		}

	$anchor{'mobility_group'} = $anchor{'city'} . '-' . $anchor{'site_code'};
	$anchor{'wlc_sysloc'} =
	  $anchor{'city'} . '_' . $anchor{'state'} . '-' . $anchor{'site_code'};

	# City name may contain spaces - remove them
	$anchor{'mobility_group'} =~ s/\s//g;
	$anchor{'wlc_sysloc'} =~ s/\s//g;

	#9800 Config
	my $wlc_nmbr = $anchor{'wlc_nmbr'};
	if ($anchor{'wlc_model'} eq "9800"){
		if ($sitetype =~ /^(T|G)$/){
			$anchor{'wlc_uplink_9800'} = smbReadFile("New-Standards/Wireless/wlc_uplink_9800_TG.txt");
			$anchor{'wlc_uplink_sec_9800'} = smbReadFile("New-Standards/Wireless/wlc_uplink_sec_9800_TG.txt");
			writeTemplate( "New-Standards/Wireless/CTRL_9800_New_Site_Model.txt",   $anchor{'wlc1_name'} . '-9800.txt' );
			writeTemplate( "New-Standards/Wireless/CTRL_9800_sec_New_Site_Model.txt",   $anchor{'wlc2_name'} . '-9800.txt' );
		}else{
			if ($wlc_nmbr == 1){
				$anchor{'wlc_uplink_9800'} = smbReadFile("New-Standards/Wireless/wlc_uplink_9800_KH.txt");
				writeTemplate( "New-Standards/Wireless/CTRL_9800_New_Site_Model.txt",   $anchor{'wlc1_name'} . '-9800.txt' );
			}else{
				$anchor{'wlc_uplink_9800'} = smbReadFile("New-Standards/Wireless/wlc_uplink_9800_KH.txt");
				$anchor{'wlc_uplink_sec_9800'} = smbReadFile("New-Standards/Wireless/wlc_uplink_sec_9800_KH.txt");
				writeTemplate( "New-Standards/Wireless/CTRL_9800_New_Site_Model.txt",   $anchor{'wlc1_name'} . '-9800.txt' );
				writeTemplate( "New-Standards/Wireless/CTRL_9800_sec_New_Site_Model.txt",   $anchor{'wlc2_name'} . '-9800.txt' );
			}
		}
	}else{
		writeTemplate( "New-Standards/Wireless/CTRL_5500.txt", $anchor{'wlc1_name'} . '.txt' );
		writeTemplate( "New-Standards/Wireless/CTRL_Initial_Config.txt", $anchor{'wlc1_name'} . '-initial.txt' );
	}
}

# VLAN definitions
sub mls_vddNew {
	( my $mls, my $sitetype ) = @_;
	my ( @dataoffset, @eacoffset, @eaclastoct, @eacgwyoct, @datapreempt, @eacpreempt );

	# Getting octect values
	my $oa = $anchor{'loop_oct1'};
	my $ob = $anchor{'loop_oct2'};
	my $dos = $anchor{'loop_oct3'};
	my $eos = 0;     						# Needed for /27 loops
	my $eacgwy = 0; 						# Needed for /27 loops
	my $eacmls = 0; 						# Needed for /27 loops

	if ( $sitetype eq 'T' ) {
		my @idfrange = ( 2..17 );           # Represents 16 IDFs

		# Loop for /24 subnets
		foreach my $offset ( @idfrange ) {
			$dos += $offset; push @dataoffset, "$oa.$ob.$dos";
			$dos -= $offset; } $dos += 18;  # Sets starting 3rd octect for EAC Loop

		# Loop for /27 subnets
		foreach my $offset ( @idfrange ) {
			push @eacoffset, "$oa.$ob.$dos"; $eacgwy = $eos + 1;
			push @eacgwyoct, $eacgwy; $eos += 32;
			if ( $eos eq 256 ) { $eos = 0; $dos++; } }

	} elsif ( $sitetype eq 'G' ) {
		$dos++; my @idfrange = ( 2..9 );     # Represents 8 IDFs

		# Loop for /27 subnets
		foreach my $offset ( @idfrange ) {
			push @eacoffset, "$oa.$ob.$dos"; $eacgwy = $eos + 1;
			push @eacgwyoct, $eacgwy; $eos += 32;
		} $dos--;

		# Loop for /24 subnets
		foreach my $offset ( @idfrange ) {
			$dos += $offset; push @dataoffset, "$oa.$ob.$dos";
			$dos -= $offset; }

	} elsif ( $sitetype eq 'K' ) {
		my @idfrange = ( 2..5 );  			# Represents 4 IDFs
		$eos = 128;	                        # Starting 4th octect for model K EAC subnets

		# Loop for /27 subnets
		foreach my $offset ( @idfrange ) {
			push @eacoffset, "$oa.$ob.$dos"; $eacgwy = $eos + 1;
			push @eacgwyoct, $eacgwy; $eos += 32; }

		# Loop for /24 subnets
		foreach my $offset ( @idfrange ) {
			$dos += $offset; push @dataoffset, "$oa.$ob.$dos";
			$dos -= $offset; }

	} elsif ( $sitetype eq 'H' ) {
		my @idfrange = ( 2,3 );				# Represents 2 IDFs
		$eos = 192; 						# Starting 4th octect for model H EAC subnets

		# Loop for /27 subnets
		foreach my $offset ( @idfrange ) {
			push @eacoffset, "$oa.$ob.$dos"; $eacgwy = $eos + 1;
			push @eacgwyoct, $eacgwy; $eos += 32; }

		# Loop for /24 subnets
		foreach my $offset ( @idfrange ) {
			$dos += $offset; push @dataoffset, "$oa.$ob.$dos";
			$dos -= $offset; }
	}

	# Defaults
	my @datavlanid   = ( 201 .. 216 );
	my @eacvlanid    = ( 401 .. 416 );
	my @datalastoct  = ( 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2,
						3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2, 3, 2 );
	my @vlanpri      = ( 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110,
						90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110, 90, 110 );

	# This is needed because the 4th octect is no longer just 3s and 2s for eac vlans
	my $rmdr = 0;
	foreach my $gwy ( @eacgwyoct ) {
		$rmdr++;
		if ( $rmdr%2 == 0 ) {						   # Condition being tested means $rmdr is an even number
			push @eaclastoct, $gwy + 2 if ( $mls == 1 );
			push @eaclastoct, $gwy + 1 if ( $mls == 2 );
		} else { push @eaclastoct, $gwy + $mls; }
	}

	if ( $mls eq '1' and $sitetype =~ /^(T|G)$/ ) {    # winds up being the reverse of the above arrays
		@datalastoct  = reverse(@datalastoct);
		@vlanpri      = reverse(@vlanpri);
	}

	# Preempts
	if ( $mls eq '1' ) {
		@datapreempt = (
						 "standby 201 preempt\r\n ", "", "standby 203 preempt\r\n ", "", "standby 205 preempt\r\n ", "",
						 "standby 207 preempt\r\n ", "", "standby 209 preempt\r\n ", "", "standby 211 preempt\r\n ", "",
						 "standby 213 preempt\r\n ", "", "standby 215 preempt\r\n "
		);
		@eacpreempt = (
						"standby 61 preempt\r\n ", "", "standby 63 preempt\r\n ", "", "standby 65 preempt\r\n ", "",
						"standby 67 preempt\r\n ", "", "standby 69 preempt\r\n ", "", "standby 71 preempt\r\n ", "",
						"standby 73 preempt\r\n ", "", "standby 75 preempt\r\n "
		);

	} else {
		@datapreempt = (
						 "", "standby 202 preempt\r\n ", "", "standby 204 preempt\r\n ", "", "standby 206 preempt\r\n ",
						 "", "standby 208 preempt\r\n ", "", "standby 210 preempt\r\n ", "", "standby 212 preempt\r\n ",
						 "", "standby 214 preempt\r\n ", "", "standby 216 preempt\r\n "
		);
		@eacpreempt = (
						"", "standby 62 preempt\r\n ", "", "standby 64 preempt\r\n ", "", "standby 66 preempt\r\n ",
						"", "standby 68 preempt\r\n ", "", "standby 70 preempt\r\n ", "", "standby 72 preempt\r\n ",
						"", "standby 74 preempt\r\n ", "", "standby 76 preempt\r\n "
		);
	}

	# Add definitions
	my $vlData = '';
	my $vlDef  = '';
	my $vlTemplate;
	my $idfct;

	# Unset currentstack so dynamic values can be used in the template
	my $tmpAnchor = delete( $anchor{'currentstack'} );

	# Data VLAN interfaces
	$vlTemplate = smbReadFile("New-Standards/Model-$sitegroup/mls_vdd_data_$sitegroup.txt");

	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $vlTmp = $vlTemplate;
		my $stack = ( split( /,/, $StackList[$ct] ) )[0];
		$vlTmp =~ s/\!vlanid\!/$datavlanid[$ct]/g;
		$vlTmp =~ s/\!currentstack\!/$stack/g;
		$vlTmp =~ s/\!data_subnet\!/$dataoffset[$ct]/g;
		$vlTmp =~ s/\!lastoct\!/$datalastoct[$ct]/g;
		$vlTmp =~ s/\!vlanpri\!/$vlanpri[$ct]/g;
		$vlTmp =~ s/\!prempt\!/$datapreempt[$ct]/g;
		$vlTmp =~ s/\!dhcp_host\!/$anchor{'dhcp_host'}/g;
		$vlTmp =~ s/\!dhcp_host_nexus\!/$anchor{'dhcp_host_nexus'}/g;
		$vlDef .= "$vlTmp\r\n";
	}

	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		( my $stack, my $switchct ) = split( /,/, $StackList[$ct] );

		# These are used in templates that are read in other parts of the script
		$anchor{'currentstack'}         = $stack;
		$anchor{'current_data_vlan'}    = $datavlanid[$ct];
		$anchor{'current_eac_vlan'}     = $eacvlanid[$ct];
		$anchor{'current_data_subnet'}  = $dataoffset[$ct];
		$anchor{'current_eac_subnet'}   = $eacoffset[$ct];

		# These are for the Visio answer file
		$idfct = $ct + 1;
		my $k = "idf$idfct";    # keys for the below
		$anchor{$k . '_flr'} = substr( $stack, 8,  3 );
		$anchor{$k . '_rm'}  = substr( $stack, 11, 3 );
		$anchor{$k . '_ds_1'}           = $dataoffset[$ct];
		$anchor{$k . '_es_1'}           = $eacoffset[$ct];
		$anchor{$k . '_Data_vlan_1'}    = $datavlanid[$ct];
		$anchor{$k . '_EAC_vlan_1'}     = $eacvlanid[$ct];
		$anchor{$k . "_dlo_$mls"} = $datalastoct[$ct];
		$anchor{$k . "_elo_$mls"} = $eaclastoct[$ct];
		$anchor{$k . '_ego'} = $eacgwyoct[$ct];
		$anchor{$k . '_eno'} = $eacgwyoct[$ct] - 1;


		if ( $mls eq '1' ) {
			writeTemplate( "New-Standards/Model-$sitegroup/stk_$switchct" . "_$sitegroup.txt", "$stack.txt" );
			writeTemplate( "New-Standards/Model-$sitegroup/aruba/stk_$switchct" . "_$sitegroup.txt", "$stack.txt" )
			if ($anchor{'stack_vendor'} eq 'aruba');
		}
	}

	# Unset currentstack (again, sheesh) so dynamic values can be used in the template
	$tmpAnchor = delete( $anchor{'currentstack'} );

	# EAC VLAN interfaces
	$vlTemplate = smbReadFile("New-Standards/Model-$sitegroup/mls_vdd_eac_$sitegroup.txt");

	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $stack      = ( split( /,/, $StackList[$ct] ) )[0];
		my $vlTmp      = $vlTemplate;
		my $eacstandby = $ct + 61;
		$vlTmp =~ s/\!vlanid\!/$eacvlanid[$ct]/g;
		$vlTmp =~ s/\!currentstack\!/$stack/g;
		$vlTmp =~ s/\!eac_subnet\!/$eacoffset[$ct]/g;
		$vlTmp =~ s/\!lastoct\!/$eaclastoct[$ct]/g;
		$vlTmp =~ s/\!gwyoct\!/$eacgwyoct[$ct]/g;
		$vlTmp =~ s/\!vlanpri\!/$vlanpri[$ct]/g;
		$vlTmp =~ s/\!prempt\!/$eacpreempt[$ct]/g;
		$vlTmp =~ s/\!dhcp_host\!/$anchor{'dhcp_host'}/g;
		$vlTmp =~ s/\!dhcp_host_nexus\!/$anchor{'dhcp_host_nexus'}/g;
		$vlTmp =~ s/\!standby_eac\!/$eacstandby/g;
		$vlDef .= "$vlTmp\r\n";
	}

	# The answer file requires blank entries for vlans that aren't configured, so create them if necessary
	$idfct++;    # increment it so we're at the next unused value
	for ( ; $idfct <= 16 ; $idfct++ ) {
		$anchor{ 'idf' . $idfct . '_flr' }  = '';
		$anchor{ 'idf' . $idfct . '_ds_1' } = '';
		$anchor{ 'idf' . $idfct . '_vs_1' } = '';
	}

	# Replace the value that was deleted just prior to the above loop for the EAC VLANs.
	$anchor{'currentstack'} = $tmpAnchor;
	return $vlDef;

}

sub mls_wddNew {
	( my $mlswireless, my $sitetype ) = @_;
	my $wddDef = '';
	if ( $sitetype =~ /^(K|H)$/ ) {
		if ( $mlswireless == 1 and $anchor{'wlc_model'} eq '9800'){
			$wddDef = smbReadFile("New-Standards/Wireless/MLS1_wireless_KH_9800.txt");
		}elsif( $mlswireless == 2 and $anchor{'wlc_model'} eq '9800'){
			$wddDef = smbReadFile("New-Standards/Wireless/MLS1_wireless_KH_with_sec_9800.txt.txt");
		}else{
			$wddDef = smbReadFile("New-Standards/Wireless/MLS1_wireless_KH.txt");
		}
	} else {
			if ( $mlswireless == 1 ){
				if ($anchor{'wlc_model'} eq '9800'){
					$wddDef = smbReadFile("New-Standards/Wireless/MLS1_wireless_9800.txt");
				}else{
					$wddDef = smbReadFile("New-Standards/Wireless/MLS1_wireless.txt");
				}
			}elsif( $mlswireless == 2 ){
				if ($anchor{'wlc_model'} eq '9800'){
					$wddDef = smbReadFile("New-Standards/Wireless/MLS2_wireless_9800.txt");
				}else{
					$wddDef = smbReadFile("New-Standards/Wireless/MLS2_wireless.txt");
				}
		 }
	}
	return $wddDef;
}

sub mls_iddNew {
	my $sitetype = shift;
	my ( @intconf, @intconf1, @intconf2, @trunkvlans, @portchannel, @portchannel_KH, );
	if ( $sitetype =~ /^(K|H)$/ ) {
		@intconf  = ( '1/1/1', '1/1/2', '1/1/3', '1/1/4' );
		@intconf1 = ( '1/1/1', '1/1/2', '1/1/3', '1/1/4' );
		@intconf2 = ( '2/1/1', '2/1/2', '2/1/3', '2/1/4' );
	} elsif ( $sitetype =~ /^(T|G)$/ ) {
		@intconf = (
					 '1/0/1',  '1/0/2',  '1/0/3',  '1/0/4',  '1/0/5',  '1/0/6',  '1/0/7',  '1/0/8',
					 '1/0/9',  '1/0/10', '1/0/11', '1/0/12', '1/0/13', '1/0/14', '1/0/15', '1/0/16',
					 '1/0/17', '1/0/18', '1/0/19', '1/0/20', '1/0/21', '1/0/22'
		);
	}
	for ( my $ct = 201 ; $ct <= 216 ; $ct++ ) {
		my $c3 = $ct + 200;
		push @trunkvlans,  $ct . ',' . $c3;    # eg '201,301,401'
		push @portchannel, $ct;
	}
	@portchannel_KH = ( 1..4 );
	my $intlimit;
	$intlimit = 2  if ( $sitetype eq 'H' );
	$intlimit = 4  if ( $sitetype eq 'K' );
	$intlimit = 8  if ( $sitetype eq 'G' );
	$intlimit = 16 if ( $sitetype eq 'T' );

	# read generic interface def file and add an interface for each stack
	my $iddOrig = smbReadFile( "New-Standards/Model-$sitegroup/mls_idd_$sitegroup.txt" );
	my $iddDef  = '';
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $iddTmp = $iddOrig;                            #make a copy of the original since symbols will be replaced multiple times
		my $idfct  = $ct + 1;
		my $stack  = ( split( /,/, $StackList[$ct] ) )[0];
		$iddTmp =~ s/\!intconf\!/$intconf[$ct]/g;
		$iddTmp =~ s/\!intconf1\!/$intconf1[$ct]/g;
		$iddTmp =~ s/\!intconf2\!/$intconf2[$ct]/g;
		$iddTmp =~ s/\!currentstack\!/$stack/g;
		$iddTmp =~ s/\!trunkvlans\!/$trunkvlans[$ct]/g;
		$iddTmp =~ s/\!portchannel\!/$portchannel[$ct]/g;
		$iddTmp =~ s/\!port-channel-m\!/$portchannel_KH[$ct]/g;
		$iddTmp .= "\r\n";
		$iddDef .= $iddTmp;

		# Save values for Visio template
		if ( $sitetype =~ /^(T|G)$/ ) {
			$anchor{ 'idf' . $idfct . '_mls_up' }   = $intconf[$ct];
			$anchor{ 'idf' . $idfct . '-1_mls_up' } = $intconf[$ct];
			$anchor{ 'idf' . $idfct . '-2_mls_up' } = $intconf[$ct];
		} elsif ( $sitetype =~ /^(K|H)$/ ) {
			$anchor{ 'idf' . $idfct . '_mls_up' }   = $intconf[$ct];
			$anchor{ 'idf' . $idfct . '-1_mls_up' } = $intconf1[$ct];
			$anchor{ 'idf' . $idfct . '-2_mls_up' } = $intconf2[$ct];
		}
	}

	# Fill in the rest of the available interfaces with a shutdown statement
	for ( my $ct = scalar(@StackList) ; $ct < $intlimit ; $ct++ ) {
		if ( $sitetype =~ /^(K|H)$/ ) {
			$iddDef .= 'interface GigabitEthernet' . $intconf[$ct] . "\r\n";
			$iddDef .= " shutdown\r\n!\r\n";
		}
	}
	return $iddDef;
}

sub cis_man_link {
	( my $sitetype, my $routerNum, my $wanlink ) = @_;

	$wanlink =~ tr/A-Z/a-z/;
	$wanlink .= '_ASR' if ( $sitetype eq 'TGKH' and $anchor{'router_type'} eq 'ASR' );

	my $readFile = "New-Standards/Model-$sitegroup/cis$routerNum". '_' ."$wanlink.txt";

	if ( $routerNum > 0 ) {
		prtout( "Writing CIS$routerNum" . ' MAN Configuration' );
	} else {
		prtout("Could not identify router number for MAN configuration");
	}
	my $wanDef = smbReadFile($readFile);
	return $wanDef;
}

sub switch_upl {
	( my $mls, my $sitetype ) = @_;

	my $uplDef = '';
	my $readFile;
	if ( $sitetype =~ /^(D|C|N|F)$/ ) {
		$mls = 2 if ( $mls > 1 );    # I think this just determines which template is read
		$readFile = "New-Standards/Model-$sitegroup/stk$mls" . "_stl_upl_$sitetype.txt";
		$readFile = "New-Standards/Model-$sitegroup/aruba/stk$mls" . "_stl_upl_$sitetype.txt"
		if ($anchor{'stack_vendor'} eq 'aruba');

	} else {
		$readFile = "New-Standards/Model-$sitegroup/mls$mls" . "_upl_$sitegroup.txt";
		$readFile = "New-Standards/Model-$sitegroup/mls$mls" . "_upl_$sitetype.txt"
		if ( $sitetype =~ /^(T|G)$/ );
	}
	$uplDef = smbReadFile($readFile);
	return $uplDef;
}

# Generate data and EAC vlan names for each stack
sub mls_vndNew {
	my $vlanDef = '';

	# Data
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $vl    = 200 + $ct + 1;
		my $datac = $ct + 1;
		$vlanDef .= 'vlan ' . $vl . "\r\n name IDF" . $datac . "_Data\r\n!\r\n";
	}

	# EAC
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $vl    = 400 + $ct + 1;
		my $datac = $ct + 1;
		$vlanDef .= 'vlan ' . $vl . "\r\n name IDF" . $datac . "_EAC\r\n!\r\n";
	}
	chop($vlanDef);    #zzz why?
	return $vlanDef;
}

# Routine to output site data
sub writeSiteNew {
		my $stacklimit = shift || 1;    # default, some site types will pass higher values

		# Site specific info
		my $sitetype = $anchor{'site_type'};
		my $stack    = substr( $anchor{'cis1_name'}, -11, 11 );
		my $wantype  = $anchor{'pri_circuit_type'};

		$anchor{'acl_tad'} = smbReadFile("New-Standards/Modules/acl_tad.txt");
		$anchor{'acl_tad'} = smbReadFile("New-Standards/Modules/aruba_acl_tad.txt")
		if ( $anchor{'stack_vendor'} eq 'aruba' );

		if (($sitetype =~ /^(D|C|N|F)$/ ) and ($anchor{'proj_type'} eq 'build')) {
			# Some of the templates have !currentstack! symbols, but that anchor value is not set
			# for some sites, so it needs to be set here.
			# Note: T, G, K and H sites set 'currentstack' in the mls_vdd sub
			# D, C, N, and F sites have only one stack so use the existing stack value
			$anchor{'currentstack'} = ( split( /,/, $StackList[0] ) )[0];
			( $stack, my $switchtype ) = split( /,/, $StackList[0] );
			$anchor{'stk_interface_uplink'} = switch_upl( $switchtype, $sitetype );

			# Hard coding the current_data_vlan in for D and N sites
			$anchor{'current_data_vlan'} = '201';

			wirelessAPNew() if ( $anchor{'proj_type'} eq 'build' );
			#VGC configs
			if ( $anchor{'vgcount'} > 0 and $anchor{'proj_type'} eq 'build' ) {
				$anchor{'vgc1_name'} = 'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01';
				$anchor{'vgc1_address'} = $anchor{'loop_subnet'} . '.134' if ( $sitetype =~ /^(N|F)$/ );
				$anchor{'vgc1_address'} = $anchor{'data_subnet_1'} . '.16' if ( $sitetype =~ /^(D|C)$/ );
				$anchor{'vgc_gwy'} = $anchor{'data_stk_oct'};
				$anchor{'vgc_mask'} = $anchor{'data_stk_mask'};
				my $of = $anchor{'vgc1_name'} . '.txt';
				$anchor{'vgc_interface_uplink'} = smbReadFile("New-Standards/Model-$sitegroup/stk1_stl_vgc.txt");
				writeTemplate( "New-Standards/Misc/vgc_400_MDF_mls1.txt", $of );

				if ( $anchor{'vgcount'} == 2 ) {
					$anchor{'vgc2_name'} = 'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a02';
					$anchor{'vgc2_address'} = $anchor{'loop_subnet'} . '.135' if ( $sitetype eq 'N' );
					$anchor{'vgc2_address'} = $anchor{'data_subnet_1'} . '.17' if ( $sitetype eq 'D' );
					my $of2 = $anchor{'vgc2_name'} . '.txt';
					$anchor{'vgc_interface_uplink'} = smbReadFile("New-Standards/Model-$sitegroup/stk2_stl_vgc2.txt");
					writeTemplate( "New-Standards/Misc/vgc_400_MDF_mls2.txt", $of2 );
				}
			} else { $anchor{'vgc_interface_uplink'} = '!'; }

			# Determine stack template based on switch type
			my %stackTemplate =
			( '1', 'stk_48p_1.txt', '2', 'stk_48p_2.txt', '3', 'stk_48p_3.txt', '4', 'stk_48p_4.txt' );
			if ( defined $stackTemplate{$switchtype} ) {
					writeTemplate( "New-Standards/Model-$sitegroup/$stackTemplate{$switchtype}", $stack . '.txt' );
					if ($anchor{'stack_vendor'} eq 'aruba'){
						if (($anchor{'primary_bu'} eq 'Genoa') and ($sitetype =~ /^(C|F)$/)){
							writeTemplate( "New-Standards/Model-$sitegroup/aruba/special_configs/Genoa/$stackTemplate{$switchtype}", $stack . '.txt' );
						}else{
							writeTemplate( "New-Standards/Model-$sitegroup/aruba/$stackTemplate{$switchtype}", $stack . '.txt' );
						}
					}
			} else {
				prtout( "Error: Switch type '$switchtype' is not a supported switch count for a $sitetype model site" );
				xit(1);
			}
			$anchor{'tad_hosts'} = smbReadFile("New-Standards/Misc/dcnf_tad_hosts.txt");

		} elsif (( $sitetype =~ /^(T|G|K|H)$/ ) and ($anchor{'proj_type'} eq 'build')) {


			if ( scalar(@StackList) > $stacklimit ) {
				prtout( "Too many stacks indentified for this site type. Please review input form.",
						"# stacks: " . scalar(@StackList) );
				xit(1);
			}

			#if( $anchor{'proj_type'} eq 'build'){
				wirelessControllerNew();
				wirelessAPNew();
				prtout("Wireless Configuration Complete");
				$anchor{'vlan_naming_dynamic'}      = mls_vndNew();
				$anchor{'interface_define_dynamic'} = mls_iddNew($sitetype);
				$anchor{'mls1_vlan_define_dynamic'} = mls_vddNew( 1, $sitetype );
				$anchor{'mls2_vlan_define_dynamic'} = mls_vddNew( 2, $sitetype );

				if ( $anchor{'wlan'} eq 'Y' ) {
					$anchor{'mls1_wireless_dynamic'} = mls_wddNew( 1, $sitetype );
					$anchor{'mls2_wireless_dynamic'} = mls_wddNew( 2, $sitetype )
					if ( $sitetype !~ /^(K|H)$/ );

				} else {
					$anchor{'mls1_wireless_dynamic'} = '';
					$anchor{'mls2_wireless_dynamic'} = '';
				}

				$anchor{'mls1_interface_uplink'} = switch_upl( 1, $sitetype );
				$anchor{'mls2_interface_uplink'} = switch_upl( 2, $sitetype )
				if ( $sitetype =~ /^(T|G)$/ );

				#VGC configs
				if ( $anchor{'vgcount'} > 0 ) {
					$anchor{'vgc1_name'} = 'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01';
					my $of = $anchor{'vgc1_name'} . '.txt';
					$anchor{'vgc1_address'} = $anchor{'loop_subnet'} . '.43' if ( $sitetype =~ /^(K|H)$/ );
					$anchor{'vgc1_address'} = $anchor{'loop_subnet'} . '.75' if ( $sitetype eq 'G' );
					$anchor{'vgc1_address'} = $anchor{'loop_subnet'} . '.139' if ( $sitetype eq 'T' );
					$anchor{'vgc_gwy'} = $anchor{'svr_gwy'};
					$anchor{'vgc_mask'} = $anchor{'svr_mask'};

					if ( $sitetype =~ /^(K|H)$/ ) {
						$anchor{'vgc_interface_uplink'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_upl_vgc1.txt");
					} elsif ( $sitetype =~ /^(T|G)$/ ){
						$anchor{'mls1_vgc_uplink'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_upl_vgc.txt");
						$anchor{'mls2_vgc_uplink'} = '!';
					} writeTemplate( "New-Standards/Misc/vgc_400_MDF_mls1.txt", $of );

					if ( $anchor{'vgcount'} == 2 ) {
						$anchor{'vgc2_name'} = 'vgc' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a02';
						my $of = $anchor{'vgc2_name'} . '.txt';
						$anchor{'vgc2_address'} = $anchor{'loop_subnet'} . '.44' if ( $sitetype =~ /^(K|H)$/ );
						$anchor{'vgc2_address'} = $anchor{'loop_subnet'} . '.76' if ( $sitetype eq 'G' );
						$anchor{'vgc2_address'} = $anchor{'loop_subnet'} . '.140' if ( $sitetype eq 'T' );

						if ( $sitetype =~ /^(K|H)$/ ){
							$anchor{'vgc_interface_uplink'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_upl_vgc2.txt");
						} elsif ( $sitetype =~ /^(T|G)$/ ) {
							$anchor{'mls1_vgc_uplink'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_upl_vgc.txt");
							$anchor{'mls2_vgc_uplink'} = smbReadFile("New-Standards/Model-$sitegroup/mls2_upl_vgc.txt");
						} writeTemplate( "New-Standards/Misc/vgc_400_MDF_mls2.txt", $of );

					} else {
						$anchor{'mls2_vgc_uplink'} = '!' ;
					}
				} else {
						if ( $sitetype =~ /^(K|H)$/ ){
							$anchor{'vgc_interface_uplink'} = '!';
						} elsif ( $sitetype =~ /^(T|G)$/ ) {
							$anchor{'mls1_vgc_uplink'} = '!';
							$anchor{'mls2_vgc_uplink'} = '!';
						}
				}

				# Firewall configs
				if ( $anchor{'fw'} eq 'Y' ) {
					$anchor{'mls1_fw_config'} = smbReadFile("New-Standards/Model-$sitegroup/mls1_fw_$sitegroup.txt");
					$anchor{'mls2_fw_config'} = smbReadFile("New-Standards/Model-$sitegroup/mls2_fw_$sitegroup.txt")
					if ( $sitetype =~ /^(T|G)$/ );
				}

				#SBC STK configs
				if ($anchor{'sbc'} eq 'yes'){
					writeTemplate( "New-Standards/Modules/stk_sbc.txt", 'stk' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01-sbc.txt' );
				}
				
				# MLS and Stack configs
				prtout("Writing MLS Configurations");
				my $template = "New-Standards/Model-$sitegroup/mls1_$sitegroup.txt";
				writeTemplate( $template, $anchor{'mls1_name'} . '.txt' );
				if ( $sitetype =~ /^(T|G)$/ ) {
					$template = "New-Standards/Model-$sitegroup/mls2_$sitegroup.txt";
					writeTemplate( $template, $anchor{'mls2_name'} . '.txt' );
				}
				$anchor{'tad_hosts'} = smbReadFile("New-Standards/Misc/tgkh_tad_hosts.txt");
			}

		if( $anchor{'proj_type'} eq 'build'){
				# TAD template (all site types use the same template)
				setTadPortsNew();    # set tad port names
			if ( $sitetype =~ /^(U)$/ ) {
				wirelessAPNew();
				$anchor{'tad_hosts'} = smbReadFile("New-Standards/Misc/4port_tad_hosts.txt");
				writeTemplate( "New-Standards/Misc/tad_u.txt", 'tad' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01_EMG8500.txt' );
			}	else {
				writeTemplate( "New-Standards/Misc/tad_g526.txt", 'tad' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01_G526.txt' );
				writeTemplate( "New-Standards/Misc/g526.txt", 'lte' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01.txt' );
				if ( $sitetype =~ /^(C|F|D|N)$/ ) {
					$anchor{'tad_hosts'} = smbReadFile("New-Standards/Misc/4port_tad_hosts.txt");
					writeTemplate( "New-Standards/Misc/tad_emg8500.txt", 'tad' . $anchor{'site_code'} . $anchor{'mdf_flr'} . 'a01_EMG8500.txt' );
				}
			}

			if ( $sitetype =~ /^(U)$/ ) {
				writeVisioU( $sitetype, 'Normal' );
				writeVisioU( $sitetype, 'Clean' );
			}	elsif ( $sitetype =~ /^(C|F)$/ ) {
				writeVisioCF( $sitetype, 'Normal' );
				writeVisioCF( $sitetype, 'Clean' );
			} elsif ( $sitetype =~ /^(D|N)$/ ) {

				# The 'overtab' data needs to be carried from the normal to the clean tab
				writeVisioDN( $sitetype, 'Normal' );
				writeVisioDN( $sitetype, 'Clean' );
			} elsif ( $sitetype =~ /^(T|G|K|H)$/ ) {
				my %stacklimit = ( 'H', 2, 'K', 4, 'G', 8, 'T', 16 );    # these site types need the stack limit to create the Visio

				# For whatever reason the 'biotab' state needs to be passed to the clean version, then added to again
				writeVisioTGKH( $sitetype, $stacklimit{$sitetype}, 'Normal' );
				writeVisioTGKH( $sitetype, $stacklimit{$sitetype}, 'Clean' );
			}
			prtout("Writing IP Summary XLS");
			my $ipSummaryFile;
			$ipSummaryFile = writeIPSummaryNew( $anchor{'site_code'}, $sitetype, $stacklimit );
			writeCISummaryNew( $anchor{'site_code'}, $sitetype, $stacklimit );
			writeEquipmentValidationNew($anchor{'site_code'});
			writeRemoteSiteBuildChecklist($anchor{'site_code'});
			if ( $wantype eq 'Metro_Ethernet' ) {

				$anchor{'cis_wan_config'} = cis_man_link( $sitetype, 1, $wantype );
				writeTemplate( "New-Standards/Model-$sitegroup/cis1_base_metroE.txt", $anchor{'cis1_name'} . '.txt' );

				$anchor{'cis_wan_config'} = cis_man_link( $sitetype, 2, $wantype )
				if ( $sitetype !~ /^(C|F|U)$/ );
				writeTemplate( "New-Standards/Model-$sitegroup/cis2_base_metroE.txt", $anchor{'cis2_name'} . '.txt' )
				if ( $sitetype !~ /^(C|F|U)$/ );

			} else {
				writeSDWANcsvNew($anchor{'site_code'},$anchor{'cis1_name'},$anchor{'cis2_name'},$anchor{'router_seltype'}, $anchor{'int_type'}, $anchor{'int_type_r1'}, $anchor{'int_type_r2'}, $anchor{'transport'});
			}
			unlink("$ROOTDIR/Files/$ipSummaryFile");
		}
		#If project type is SDWAN ONLY
		elsif($anchor{'proj_type'} eq 'proj-sdwan'){
			#Force SDWAN anchor to YES
			$anchor{'SDWAN'} = 'Y';
			writeSDWANcsvNew($anchor{'site_code'},$anchor{'cis1_name'},$anchor{'cis2_name'},$anchor{'router_seltype'}, $anchor{'int_type'}, $anchor{'int_type_r1'}, $anchor{'int_type_r2'}, $anchor{'transport'});
		}

		my $zipfile = compress();

		prtout( "Configurator Output Generation Complete.<br/>",
				"<a HREF='/tmp/$OutputDir.zip' >D&E and IP Summary can be found here</a>" );

				# #UPS stk connections
				# $anchor{'ups_currentstack'} = substr($anchor{'currentstack'}, 8, 4);
				# prtout"UPS stk on global: $anchor{'ups_currentstack'}";
}

sub writeVisioU {

	return unless ($anchor{'proj_type'} eq 'build');
		$anchor{'street'} =~s/&/&amp;/g;

	( my $sitetype, my $vistype ) = @_;
	my $sitegroupTmp = 'u';

	# Set the template and output file
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-" . $sitegroupTmp . '-template.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-" . $sitegroupTmp . '-template-clean.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';
	}

	my $backbiotab   = 0;
	my $backmdftab   = 4;
	my $mdfutab     = 5;
	my $backipttab   = 6;
	my $biographytab = 7;
	my $cistab       = 8;    # generic for U sites
	my $ipttab       = 10;
	my $tadtab  	 = 13;
	# my $stkovertab   = 16;

	prtout("Opening Visio Template for Processing");
	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );    #zzz error handling
	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);

	# Upload tabs
	my %tabs = VisioReadTabs( \$fill );

	my $mdftab =
	    $tabs{$biographytab}
	  . $tabs{$mdfutab}
	  . $tabs{$ipttab}
	  . $tabs{$backbiotab}
	  . $tabs{$backmdftab}
	  . $tabs{$backipttab};

	#router
	VisioControlLayer( 'C1111X-8P',                 0, \$mdftab );
	VisioControlLayer( 'C1161X-8P',                0, \$mdftab );
	VisioControlLayer( 'C1121X-8P',                0, \$mdftab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$mdftab );

	my $overview = $tabs{$cistab};
	VisioControlLayer( 'C1111X-8P',                 0, \$overview );
	VisioControlLayer( 'C1161X-8P',                0, \$overview );
		VisioControlLayer( 'C1121X-8P',                0, \$overview );
	VisioControlLayer( $anchor{'router_type'}, 1, \$overview );

	# SDWAN
	VisioControlLayer( 'SDWAN', 0, \$mdftab ) if ($anchor{'SDWAN'} ne 'Y');
	VisioControlLayer( 'NON-SDWAN', 0, \$mdftab ) if ($anchor{'SDWAN'} eq 'Y');
	VisioControlLayer( '2_transports', 0, \$mdftab ) if ($anchor{'transport'} == 1 ); # Turn off 2_transports layer when there is only 1 circuit

	# Cellular Gateways
	if ($anchor{'isp'} ne 'CELLULAR' ) { # If a CG is not set to primary
		VisioControlLayer( 'cellular_pri', 0, \$mdftab );
		VisioControlLayer( 'cellular_dual', 0, \$mdftab );
	}
	if ($anchor{'pri_provider'} ne 'CELLULAR' ) { # If a CG is not set to secondary
		VisioControlLayer( 'cellular_sec', 0, \$mdftab );
		VisioControlLayer( 'cellular_dual', 0, \$mdftab );
	}
	if ($anchor{'isp'} eq 'CELLULAR' and $anchor{'pri_provider'} eq 'CELLULAR' ) { # Dual CGs
		VisioControlLayer( 'cellular_sec', 0, \$mdftab );
	}

	my $tad = $tabs{$tadtab};
	VisioControlLayer( 'aruba', 0, \$tad ) if ( $anchor{'stack_vendor'} eq 'cisco' );
	VisioControlLayer( 'cisco', 0, \$tad ) if ( $anchor{'stack_vendor'} eq 'aruba' );

	# Search and replace variables
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
		$tad =~ s/\!$key\!/$anchor{$key}/g;
	}

	# Put together tabs
	my $generatedTabs = $mdftab . $tad . $overview;
	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified $vistype Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");
	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );             #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

	unlink("$ROOTDIR/Files/$vOutput");
}

sub writeVisioCF {

	return unless ($anchor{'proj_type'} eq 'build');
		$anchor{'street'} =~s/&/&amp;/g;

	( my $sitetype, my $vistype ) = @_;
	my $sitegroupTmp = 'cf';

	# Set the template and output file
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-" . $sitegroupTmp . '-template.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-" . $sitegroupTmp . '-template-clean.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';
	}

	my $backbiotab   = 0;
	my $backmdftab   = 4;
	my $mdfcftab     = 5;
	my $backipttab   = 6;
	my $biographytab = 7;
	my $cistab       = 8;    # generic for F and C sites
	my $ipttab       = 10;
	my $tadtab  	 = 13;
	my $stkovertab   = 16;

	prtout("Opening Visio Template for Processing");
	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );    #zzz error handling
	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);

	# Upload tabs
	my %tabs = VisioReadTabs( \$fill );

	my $mdftab =
	    $tabs{$biographytab}
	  . $tabs{$mdfcftab}
	  . $tabs{$ipttab}
	  . $tabs{$backbiotab}
	  . $tabs{$backmdftab}
	  . $tabs{$backipttab};

	#router
	VisioControlLayer( '4331',                 0, \$mdftab );
	VisioControlLayer( 'C8200-1N-4T',                0, \$mdftab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$mdftab );

	my $overview = $tabs{$cistab};
	VisioControlLayer( '4331',                 0, \$overview );
	VisioControlLayer( 'C8200-1N-4T',                0, \$overview );
	VisioControlLayer( $anchor{'router_type'}, 1, \$overview );

	# SDWAN
	VisioControlLayer( 'SDWAN', 0, \$mdftab ) if ($anchor{'SDWAN'} ne 'Y');
	VisioControlLayer( 'NON-SDWAN', 0, \$mdftab ) if ($anchor{'SDWAN'} eq 'Y');

	# Wireless
	VisioControlLayer( 'WLAN',    0, \$mdftab ) if ( $anchor{'wlan'} ne 'Y' );
	VisioControlLayer( 'WLANCFG', 0, \$mdftab ) if ( $anchor{'wlan'} ne 'Y' or $sitetype eq 'F' );

	my $stacktab = $tabs{$stkovertab};
	VisioControlLayer( 'aruba', 0, \$stacktab ) if ( $anchor{'stack_vendor'} eq 'cisco' );
	VisioControlLayer( 'cisco', 0, \$stacktab ) if ( $anchor{'stack_vendor'} eq 'aruba' );

	my $tad = $tabs{$tadtab};
	VisioControlLayer( 'aruba', 0, \$tad ) if ( $anchor{'stack_vendor'} eq 'cisco' );
	VisioControlLayer( 'cisco', 0, \$tad ) if ( $anchor{'stack_vendor'} eq 'aruba' );

	my $switchtype = ( split( /,/, $StackList[0] ) )[1];

	doswitchesNew( $switchtype, 1, \$mdftab );

	# Search and replace variables
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
		$tad =~ s/\!$key\!/$anchor{$key}/g;
	}

	# Logic for vgc layers
	VisioControlLayer( 'vgc_port', 0, \$mdftab )
	if ( $anchor{'stack_vendor'} eq 'aruba' or $anchor{'vgcount'} == 0 );

	VisioControlLayer( 'vgc_port_aruba', 0, \$mdftab )
	if ( $anchor{'stack_vendor'} eq 'cisco' or $anchor{'vgcount'} == 0 );

	VisioControlLayer( 'VGC', 0, \$mdftab ) unless ( $anchor{'vgcount'} > 0 );

	# Put together tabs
	my $generatedTabs = $mdftab . $stacktab . $tad . $overview;
	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified $vistype Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");
	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );             #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

	unlink("$ROOTDIR/Files/$vOutput");
}

sub writeVisioDN {
return unless ($anchor{'proj_type'} eq 'build');

	$anchor{'street'} =~s/&/&amp;/g;

	( my $sitetype, my $vistype, my $overtab ) = @_;
	$overtab = '' if ( !( defined $overtab ) );
	my $fw   = $anchor{'fw'};
	my $wlan = $anchor{'wlan'};
	my $sitegroupTmp = 'dn';

	# Set template and output file
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-" .$sitegroupTmp . '-template.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-" . $sitegroupTmp . '-template-clean.vdx';
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';

		# Set TLOC interfaces for use on the clean DNE
		$anchor{'tloc_int_clean'} = stripSubInt($anchor{'tloc_int'});
		$anchor{'tloc_int2_clean'} = stripSubInt($anchor{'tloc_int2'});
	}
	my $backmdftab   = 4;	my $mdfdntab     = 5;
	my $backipttab   = 6;	my $backbiotab   = 9;
	my $biographytab = 10;	my $tadtab   = 14;
	my $fwoverinttab = 15;	my $backsectab   = 16;
	my $secl3tab     = 17;	my $stkovertab   = 19;
	my $waetab       = 22;	my $ipttab       = 24;
	my $fwoverexttab = 26;	my $routertab    = 23;
	my $armistab 		 = 29;	my $armisovertab = 30;

	prtout("Opening Visio Template for Processing");

	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );    #zzz error handling

	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);

	# Upload tabs
	my %tabs = VisioReadTabs( \$fill );

	# Define tabs to variables
	my $biotab = $tabs{$biographytab};

	# Add TAD tab for sites with IPS
	my $mdftab = "$tabs{$mdfdntab}";

	# Security Layer3 and FW Overviews
	$overtab .= "$tabs{$secl3tab}$tabs{$fwoverinttab}" if ( $fw eq 'Y' );

	#MDF Tab FW Overview
	VisioControlLayer( 'FW',    0, \$mdftab ) if ( $fw ne 'Y' );
	VisioControlLayer( "stack1_cable_fw", 0, \$mdftab );
	VisioControlLayer( "stack2_cable_fw", 0, \$mdftab );

	# WAAS overview, IPT Template and Stack overview
	$overtab .= "$tabs{$ipttab}";

	#Armis overview tab
	$armistab .= "$tabs{$armistab}$tabs{$armisovertab}";

	# 4/20/2017 - simplified code
	#           - also using $routertab instead of router specific tabs
	VisioControlLayer( '4321',                 0, \$mdftab );
	VisioControlLayer( '4331',                 0, \$mdftab );
	VisioControlLayer( '4351',                 0, \$mdftab );
	VisioControlLayer( '4451',                 0, \$mdftab );
	VisioControlLayer( 'C8200-1N-4T',                0, \$mdftab );
	VisioControlLayer( 'C8300-1N1S',                0, \$mdftab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$mdftab );
	$overtab .= $tabs{$routertab};
	VisioControlLayer( '4321',                 0, \$overtab );
	VisioControlLayer( '4331',                 0, \$overtab );
	VisioControlLayer( '4351',                 0, \$overtab );
	VisioControlLayer( '4451',                 0, \$overtab );
	VisioControlLayer( 'C8200-1N-4T',                0, \$overtab );
	VisioControlLayer( 'C8300-1N1S',                0, \$overtab );
	VisioControlLayer( $anchor{'router_type'}, 1, \$overtab );

	my $stacktab = $tabs{$stkovertab};
	VisioControlLayer( 'aruba', 0, \$stacktab ) if ( $anchor{'stack_vendor'} eq 'cisco' );
	VisioControlLayer( 'cisco', 0, \$stacktab ) if ( $anchor{'stack_vendor'} eq 'aruba' );

	my $tad = $tabs{$tadtab};
	VisioControlLayer( 'aruba', 0, \$tad ) if ( $anchor{'stack_vendor'} eq 'cisco' );
	VisioControlLayer( 'cisco', 0, \$tad ) if ( $anchor{'stack_vendor'} eq 'aruba' );

	my $switchtype = ( split( /,/, $StackList[0] ) )[1];

	doswitchesNew( $switchtype, 1, \$mdftab );

	my $backtab = "$tabs{$backmdftab}$tabs{$backipttab}$tabs{$backbiotab}$tabs{$backsectab}";

	# Logic for vgc layers
	VisioControlLayer( 'vgc_port', 0, \$mdftab )
	if ( $anchor{'stack_vendor'} eq 'aruba' or $anchor{'vgcount'} == 0 );

	VisioControlLayer( 'vgc_port_aruba', 0, \$mdftab )
	if ( $anchor{'stack_vendor'} eq 'cisco' or $anchor{'vgcount'} == 0 );

	VisioControlLayer( 'vgc2_port', 0, \$mdftab )
	if ( $anchor{'stack_vendor'} eq 'aruba' or $anchor{'vgcount'} < 2 );

	VisioControlLayer( 'vgc2_port_aruba', 0, \$mdftab )
	if ( $anchor{'stack_vendor'} eq 'cisco' or $anchor{'vgcount'} < 2 );

	VisioControlLayer( 'VGC', 0, \$mdftab ) unless ( $anchor{'vgcount'} > 0 );
	VisioControlLayer( 'VGC2', 0, \$mdftab ) unless ( $anchor{'vgcount'} == 2 );

	# SDWAN

	VisioControlLayer( '2_transports', 0, \$mdftab )
	if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} eq 'yes');

	VisioControlLayer( '3_transports', 0, \$mdftab )
	if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} eq 'yes');


	VisioControlLayer( '2_transpo_new', 0, \$mdftab )
	if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} ne 'yes');

	VisioControlLayer( '3_transpo_new', 0, \$mdftab )
	if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} ne 'yes');

	VisioControlLayer( 'non_sdwan', 0, \$mdftab ) if ($anchor{'SDWAN'} eq 'Y');

	# Wireless
	VisioControlLayer( 'WLAN',    0, \$mdftab ) if ( $wlan ne 'Y' );
	VisioControlLayer( 'WLANCFG', 0, \$mdftab )
	if ( $anchor{'wlan'} ne 'Y' or $sitetype eq 'N' );

	# Search and replace variables
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
		$overtab =~ s/\!$key\!/$anchor{$key}/g;
		$backtab =~ s/\!$key\!/$anchor{$key}/g;
		$tad =~ s/\!$key\!/$anchor{$key}/g;
		$armistab =~ s/\!$key\!/$anchor{$key}/g;
	}

	# Put together tabs
	my $generatedTabs = $biotab . $mdftab . $overtab . $backtab . $stacktab . $tad . $armistab;
	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified $vistype Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");

	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );             #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

	unlink("$ROOTDIR/Files/$vOutput");
	return $overtab;
}

sub writeVisioTGKH {
	return unless ($anchor{'proj_type'} eq 'build');
		$anchor{'street'} =~s/^ | &//g;
		$anchor{'street'} =~s/&/&amp;/g;

	( my $sitetype, my $stacklimit, my $vistype ) = @_;
	my ( $vTemplate, $vOutput );
	if ( $vistype eq 'Normal' ) {
		$vTemplate = "$ROOTDIR/master-tgkh-template.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0.vdx';
	} else {
		$vTemplate = "$ROOTDIR/master-tgkh-template-clean.vdx";
		$vOutput   = $anchor{'site_code'} . '-dne-v1.0-clean.vdx';

		# Set TLOC interfaces for use on the clean DNE
		$anchor{'tloc_int_clean'} = stripSubInt($anchor{'tloc_int'});
		$anchor{'tloc_int2_clean'} = stripSubInt($anchor{'tloc_int2'});
	}
	my $idf2tab       = 0;   my $idf3tab       = 4;
	my $backmdftab    = 5;   my $mdfnextab     = 6;
	my $mdftgtab      = 7;   my $idf1tab       = 8;
	my $mdfkhtab      = 9;   my $backstktab    = 10;
	my $backidfxab    = 11;	 my $mdfstktab     = 12;
	my $backipttab    = 14;	 my $backbiotab    = 16;
	my $biographytab  = 17;	 my $cis3945tab    = 19;
	my $backsectab    = 24;	 my $tadtab        = 27;
	my $backwlantab   = 30;  my $wirelesstab   = 31;
	my $stkovertab    = 33;  my $waetab        = 34;
	my $ipttab        = 36;	 my $routerovertab = 45;
	my $fwovertab     = 996; my $dmzflowtab    = 997;
	my $untrusttab    = 998; my $secl3tab      = 999;
	my @idfxabs       = ();	 my $thisidf       = '';
	my $lastidf       = '';	 my $store         = '';

	# Save stack names (space separated) for each unique idf. I think. The old code does this in a weird way
	# and I might have missed something.
	my @idforder;
	my $stacklist = '';
	foreach my $stk ( @StackList, 'end' ) {    # Adding 'end' forces an update in the final iteration of the loop.
		                                       # 'end' can be anything other than a valid stack name.
		my $idf = substr( $stk, 9, 3 );
		if ( $lastidf eq '' ) {                # first stack in loop iteration
			$lastidf   = $idf;
			$stacklist = $stk;
		} elsif ( $idf ne $lastidf ) {         # When the stack name has changed,
			push @idforder, $stacklist;        # save off the stack list
			$stacklist = $stk;                 # and reset the stack list to the current stack name.
		} else {
			$stacklist .= " $stk";             # Adds stack name to existing (space delimited) list
		}
		$lastidf = $idf;
	}
	my $idfxabid = 50;
	my $idfct    = 0;
	for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
		my $stack    = $ct + 1;
		my $name     = 'sc' . $stack;
		my $stkcount = ( split( /,/, $StackList[$ct] ) )[1];
		$anchor{$name} = $stkcount;
	}
	my $fw   = $anchor{'fw'};
	my $ipt  = $anchor{'ipt'};
	my $wlan = $anchor{'wlan'};
	prtout("Opening Visio Template for Processing");
	my $fill = '';
	open( VISIO, "<:utf8", $vTemplate );

	while (<VISIO>) {
		$fill .= $_;
	}
	close(VISIO);
	my %tabs = VisioReadTabs( \$fill );

	# Add Bio Page
	my $biotab .= $tabs{$biographytab};

	# Add MDF tabs
	# show stack overview sitetype layer
	$mdfstktab = $tabs{$mdfstktab};
	VisioControlLayer( 'idftg', 0, \$mdfstktab ) if ( $sitetype !~ /^(T|G)$/ );
	VisioControlLayer( 'idfkh', 0, \$mdfstktab ) if ( $sitetype !~ /^(K|H)$/ );

	my $mdftab;
	if ( $sitetype =~ /^(K|H)$/ ) {
		$mdftab = $tabs{$mdfkhtab} . $mdfstktab;
	} elsif ( $sitetype =~ /^(T|G)$/ ) {
		$mdftab = $tabs{$mdftgtab} . $mdfstktab;
	}

	# Set all to hidden by default
	VisioControlLayer( '4331',  		0, \$mdftab );
	VisioControlLayer( '4351',  		0, \$mdftab );
	VisioControlLayer( '4451',  		0, \$mdftab );
	VisioControlLayer( '4461',  		0, \$mdftab );
	VisioControlLayer( 'ASR',   		0, \$mdftab );
	VisioControlLayer( 'C8200-1N-4T',   		0, \$mdftab );
	VisioControlLayer( 'C8300-1N1S',   	0, \$mdftab );
	VisioControlLayer( 'C8300-2N2S',   	0, \$mdftab );
	VisioControlLayer( 'C8500-12X4QC',   	0, \$mdftab );

	# Make the appropriate layer visible
	VisioControlLayer( $anchor{'router_type'}, 1, \$mdftab );


	# VGC
	VisioControlLayer( 'VGC1', 0, \$mdftab ) unless ( $anchor{'vgcount'} > 0 );
	VisioControlLayer( 'VGC2', 0, \$mdftab ) unless ( $anchor{'vgcount'} == 2 );

	# SDWAN
	if ( $sitetype =~ /^(K|H)$/ ) {

		VisioControlLayer( '2_transports', 0, \$mdftab )
		if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} eq 'yes');

		VisioControlLayer( '3_transports', 0, \$mdftab )
		if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} eq 'yes');

		VisioControlLayer( '2_transpo_new', 0, \$mdftab )
		if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} ne 'yes');

		VisioControlLayer( '3_transpo_new', 0, \$mdftab )
		if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} ne 'yes');

		VisioControlLayer( 'non_sdwan', 0, \$mdftab ) if ($anchor{'SDWAN'} eq 'Y');

	} else {

		VisioControlLayer( '2_transports', 0, \$mdftab )
		if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} eq 'yes');

		VisioControlLayer( '3_transports', 0, \$mdftab )
		if ($anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} eq 'yes');


		VisioControlLayer( '2_transpo_new', 0, \$mdftab )
		if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 2 or $anchor{'tloc'} ne 'yes');

		VisioControlLayer( '3_transpo_new', 0, \$mdftab )
		if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 3 or $anchor{'tloc'} ne 'yes');

		VisioControlLayer( '4_transpo_new', 0, \$mdftab )
		if ( $anchor{'SDWAN'} ne 'Y' or $anchor{'transport'} != 4 or $anchor{'tloc'} ne 'yes');

	}

	# Firewall
	VisioControlLayer( 'FW',    0, \$mdftab ) if ( $fw ne 'Y' );

	#Wireless
	VisioControlLayer( 'WLAN',  0, \$mdftab ) if ( $wlan ne 'Y' );

	# Layer control for stacks on MDF tab - configurator allows 16 stacks maximum
	my $stkct = scalar(@StackList);
	for ( my $ct = 16 ; $ct > 1 ; $ct-- ) {
		VisioControlLayer( 'mdf_stack' . $ct, 0, \$mdftab ) if ( $stkct < $ct );
	}

	# Search and replace values in mdf and bio tabs
	foreach my $key ( keys %anchor ) {
		$mdftab =~ s/\!$key\!/$anchor{$key}/g;
		$biotab =~ s/\!$key\!/$anchor{$key}/g;
	}
	# Creating idf tabs
	my $generatedTabs = $biotab . $mdftab . $tabs{$ipttab};
	my ( $swcount, $workingtab );

	foreach my $idf ( @idforder ) {
		my @switches = split( /\s/, $idf );
		$swcount    = scalar(@switches);
		$workingtab = $tabs{$idf1tab} if ( $swcount == 1 );    # one stack per IDF
		$workingtab = $tabs{$idf2tab} if ( $swcount == 2 );    # one stack per IDF
		$workingtab = $tabs{$idf3tab} if ( $swcount == 3 );    # one stack per IDF
		if ( scalar(@switches) > 3 ) {                         # limited to 3 to keep Visio page clean looking
			prtout( "There are too many stacks in this IDF for Configurator to handle.",
					"Please split them into two IDFs and retrun." );
			xit(1);
		}
		my $idfname =
		  substr( $switches[0], 8, 3 ) . '-' . substr( $switches[0], 11, 1 );
		VisioRenameTab( "IDF " . $idfname, \$workingtab );
		$idfxabid++;
		VisioReIDTab( $idfxabid, \$workingtab );

		# Search and replace values in idf tabs
		#zzz Cleaner way to do this, as asked on legacy section
		for ( my $peridf = 2; $peridf >= 0; $peridf-- ) {
			if ( $swcount > $peridf ) {
				my %keyrvals = ( '_ds_1', '_ds_1', '_es_1', '_es_1', '_Data_vlan_1', 'dvlan',
					'_EAC_vlan_1', 'evlan', '_dlo_1', '_dlo1', '_dlo_2', '_dlo2',
					'_elo_1', '_elo1', '_elo_2', '_elo2', '_mls_up', 'up', '-1_mls_up',
					'-1up', '-2_mls_up', '-2up', '_eno', '_eno', '_ego', '_ego' );
				my $idfct = $idfct+ $peridf + 1;
				my $val  = $peridf + 1;
				my ( $key, $rval );

				foreach my $kr ( keys %keyrvals ) {
					$key   = "idf$idfct$kr";
					$rval  = "idf$val$keyrvals{$kr}";
					$rval  = "$keyrvals{$kr}$val" if ( $kr =~ /vlan/ );
					$workingtab =~ s/\!$rval\!/$anchor{$key}/g;
				}
				my ( $stkname, $swcount ) = split( /,/, $switches[$peridf] );
				$rval = "!stack" . $val . '_name!';
				$workingtab =~ s/$rval/$stkname/g;
				doswitchesNew( $swcount, $val, \$workingtab );
			}
		}
		$workingtab =~ s/\!idf_name\!/$idfname/g;
		foreach my $key ( keys %anchor ) {
			$workingtab =~ s/\!$key\!/$anchor{$key}/g;
		}

		VisioControlLayer( 'idftg', 0, \$workingtab ) if ( $sitetype !~ /^(T|G)$/ );
		VisioControlLayer( 'idfkh', 0, \$workingtab ) if ( $sitetype !~ /^(K|H)$/ );
		$generatedTabs .= $workingtab;
		$idfct += $swcount;
	}

	# Add background tabs
	my $idfbg = $tabs{$backidfxab};

	# Disable unnecessary IDF layers
	VisioControlLayer( 'idftg', 0, \$idfbg ) if ( $sitetype !~ /^(T|G)$/ );
	VisioControlLayer( 'idfkh', 0, \$idfbg ) if ( $sitetype !~ /^(K|H)$/ );

	my $mdfbg  = $tabs{$backmdftab};
	my $stabg  = $tabs{$backstktab};
	my $iptbg  = $tabs{$backipttab};
	my $biobg  = $tabs{$backbiotab};
	my $secbg  = $tabs{$backsectab};
	my $wlanbg = $tabs{$backwlantab};

	if ($anchor{'wlc_model'} ne '9800'){
		#Make selection for 'idftg_old'
		if ($sitetype =~ /^(T|G)$/){
			VisioControlLayer( 'idftg', 0, \$wlanbg );
			VisioControlLayer( 'idfkh', 0, \$wlanbg );
			VisioControlLayer( 'idfkh_2', 0, \$wlanbg );
			VisioControlLayer( 'idfkh_old', 0, \$wlanbg );
		}
		#Make selection for 'idfkh_old'
		else{
			VisioControlLayer( 'idftg', 0, \$wlanbg );
			VisioControlLayer( 'idfkh', 0, \$wlanbg );
			VisioControlLayer( 'idfkh_2', 0, \$wlanbg );
			VisioControlLayer( 'idftg_old', 0, \$wlanbg );
		}
	}elsif ($sitetype =~ /^(K|H)$/){
		if ( ($anchor{'wlc_nmbr'} == 1) ){
		VisioControlLayer( 'idfkh_2', 0, \$wlanbg );
		VisioControlLayer( 'idftg', 0, \$wlanbg );
		VisioControlLayer( 'idftg_old', 0, \$wlanbg );
		VisioControlLayer( 'idfkh_old', 0, \$wlanbg );
	}elsif ( ($anchor{'wlc_nmbr'} == 2) ){
		VisioControlLayer( 'idfkh', 0, \$wlanbg );
		VisioControlLayer( 'idftg', 0, \$wlanbg );
		VisioControlLayer( 'idftg_old', 0, \$wlanbg );
		VisioControlLayer( 'idfkh_old', 0, \$wlanbg );
	}
	prtout( "VISIO WLC NUMBER: $anchor{'wlc_nmbr'}\n" );
}else{
	#VisioControlLayer( 'idftg', 0, \$wlanbg ) if ( $sitetype !~ /^(T|G)$/ );
	VisioControlLayer( 'idftg_old', 0, \$wlanbg );
	VisioControlLayer( 'idfkh_old', 0, \$wlanbg );
	VisioControlLayer( 'idfkh', 0, \$wlanbg );
	VisioControlLayer( 'idfkh_2', 0, \$wlanbg );
	}

	my $tad = $tabs{$tadtab};
	VisioControlLayer( 'idftg', 0, \$tad ) if ( $sitetype !~ /^(T|G)$/ );
	VisioControlLayer( 'idfkh', 0, \$tad ) if ( $sitetype !~ /^(K|H)$/ );

	# Do substitutions for these tabs
	foreach my $key ( keys %anchor ) {
		$idfbg =~ s/\!$key\!/$anchor{$key}/g;
		$mdfbg =~ s/\!$key\!/$anchor{$key}/g;
		$stabg =~ s/\!$key\!/$anchor{$key}/g;
		$iptbg =~ s/\!$key\!/$anchor{$key}/g;
		$biobg =~ s/\!$key\!/$anchor{$key}/g;
		$secbg =~ s/\!$key\!/$anchor{$key}/g;
		$wlanbg =~ s/\!$key\!/$anchor{$key}/g;
		$tad =~ s/\!$key\!/$anchor{$key}/g;
	}
	$generatedTabs .= $tad;

	# Wireless tab
	if ( $wlan eq 'Y' ) {
		my $wlantab = $tabs{$wirelesstab};
		foreach my $key ( keys %anchor ) {
			$wlantab =~ s/\!$key\!/$anchor{$key}/g;
		}
		#Make old WLC ON
		if ( $anchor{'wlc_model'} ne '9800' ){
			VisioControlLayer( 'idftg', 0, \$wlantab );
			VisioControlLayer( 'idfkh', 0, \$wlantab );
			VisioControlLayer( 'idfkh_2', 0, \$wlantab );
		}elsif ($sitetype =~ /^(K|H)$/){
			if ( ($anchor{'wlc_nmbr'} == 1) ){
				VisioControlLayer( 'idfkh_2', 0, \$wlantab );
				VisioControlLayer( 'idftg', 0, \$wlantab );
				VisioControlLayer( 'idftgkh_old', 0, \$wlantab );
				}elsif ( ($anchor{'wlc_nmbr'} == 2) ){
				VisioControlLayer( 'idfkh', 0, \$wlantab );
				VisioControlLayer( 'idftg', 0, \$wlantab );
				VisioControlLayer( 'idftgkh_old', 0, \$wlantab );
				}
		}else{
		VisioControlLayer( 'idftgkh_old', 0, \$wlantab );
		VisioControlLayer( 'idfkh', 0, \$wlantab );
		VisioControlLayer( 'idfkh_2', 0, \$wlantab );
		}

		$generatedTabs .= $wlantab;
	}
	# Firewall tabs
	if ( $fw eq 'Y' ) {
		my $secl3 = $tabs{$secl3tab};
		VisioControlLayer( 'idftg', 0, \$secl3 ) if ( $sitetype !~ /^(T|G)$/ );
		VisioControlLayer( 'idfkh', 0, \$secl3 ) if ( $sitetype !~ /^(K|H)$/ );
		my $fwovint = $tabs{$fwovertab};
		my $dmzflow = $tabs{$dmzflowtab};
		my $utdmz   = $tabs{$untrusttab};
		VisioControlLayer( 'idftg', 0, \$utdmz) if ( $sitetype !~ /^(T|G)$/ );
		VisioControlLayer( 'idfkh', 0, \$utdmz ) if ( $sitetype !~ /^(K|H)$/ );

		foreach my $key ( keys %anchor ) {
			$secl3 =~ s/\!$key\!/$anchor{$key}/g;
			$fwovint =~ s/\!$key\!/$anchor{$key}/g;
			$utdmz =~ s/\!$key\!/$anchor{$key}/g;
			$dmzflow =~ s/\!$key\!/$anchor{$key}/g;
		}
		$generatedTabs .= $secl3 . $utdmz . $fwovint . $dmzflow;
	}

	#stack overview tab
	my $stacktab = $tabs{$stkovertab};
	VisioControlLayer( 'Aruba', 0, \$stacktab ) if ( $anchor{'stack_vendor'} eq 'cisco' );
	VisioControlLayer( 'Cisco', 0, \$stacktab ) if ( $anchor{'stack_vendor'} eq 'aruba' );

	# New 4/20/2017 - Routers are combined into a single tab named 'Router Overview' that uses layers
	$routerovertab = $tabs{$routerovertab};
	VisioControlLayer( '4331', 0, \$routerovertab );
	VisioControlLayer( '4351', 0, \$routerovertab );
	VisioControlLayer( '4451', 0, \$routerovertab );
	VisioControlLayer( '4461', 0, \$routerovertab );
	VisioControlLayer( 'ASR',  0, \$routerovertab );
	VisioControlLayer( 'C8200-1N-4T',   		0, \$routerovertab );
	VisioControlLayer( 'C8300-1N1S',   	0, \$routerovertab );
	VisioControlLayer( 'C8300-2N2S',   	0, \$routerovertab );
	VisioControlLayer( 'C8500-12X4QC',   	0, \$routerovertab );

	# Make the appropriate layer visible
	VisioControlLayer( $anchor{'router_type'}, 1, \$routerovertab );

	$generatedTabs .= $routerovertab;

	$generatedTabs .= $mdfbg . $stabg . $idfbg . $iptbg . $biobg . $secbg . $wlanbg . $stacktab;
	substr( $fill, index( $fill, '</Pages>' ), 0 ) = $generatedTabs;
	prtout("Writing out modified Visio template");

	open( OUT, ">:utf8", "$SVR_ROOTDIR/$OutputDir/$vOutput" );    #zzz error handling
	print OUT $fill;
	close(OUT);

	unlink("$ROOTDIR/Files/$vOutput");

	return;

	open( OUT, ">:utf8", "$ROOTDIR/Files/$vOutput" );             #zzz error handling
	print OUT $fill;
	close(OUT);
	smbPut( "$ROOTDIR/Files/$vOutput", "$SMB_FIN_DIR/$OutputDir/$vOutput" );

 # Return this - after the 'Normal' Visio runs the 'Clean' needs to add to this value
	return $biotab;
}

sub doswitchesNew {
return unless ($anchor{'proj_type'} eq 'build');
my $fw   = $anchor{'fw'};
	( my $swct, my $chstack, my $data ) = @_;
	my $stklayer = $anchor{'stack_vendor'};

	for ( my $ct = 4; $ct >= 1; $ct-- ) {
		VisioControlLayer( "idf$chstack" . "_cisco_stk_sw$ct", 0, $data ); #idf1_cisco_stk_sw1, hide, mdftab
		VisioControlLayer( "idf$chstack" . "_aruba_stk_sw$ct", 0, $data );
		VisioControlLayer( "idf$chstack" . "_stack_cable$ct",  0, $data );
	}

	for ( my $ct = 1 ; $ct <= 4 ; $ct++ ) {
		if ( $swct >= $ct ) {
			VisioControlLayer( "idf$chstack" . "_$stklayer" . "_stk_sw$ct", 1, $data );
			VisioControlLayer( "idf$chstack" . "_stack_cable$ct",  			1, $data );
		}
	}
	VisioControlLayer( "idf$chstack" . "_stack_cable1",  0, $data ) if ( $swct > 1 );

	VisioControlLayer( "idf$chstack" . '_cisco_stk_p1', 0, $data )
	if ( $swct > 1 or $stklayer eq 'aruba' );

	VisioControlLayer( "idf$chstack" . '_aruba_stk_p1', 0, $data )
	if ( $swct > 1 or $stklayer eq 'cisco' );

}

sub setTadPortsNew {
	return unless ($anchor{'proj_type'} eq 'build');
	# Only for H, K, G and T sites
	$anchor{'tad_port_1'}  = $anchor{'cis1_name'};
	$anchor{'tad_port_2'}  = 'Optional_2';
	$anchor{'tad_port_2'}  = $anchor{'cis2_name'} if ( $sitetype !~ /^(C|F|U)$/ );
	$anchor{'tad_port_2'}  = $anchor{'cgw1_name'} if ( ( $sitetype =~ /^(U)$/ ) and ($anchor{'isp'} eq 'CELLULAR' or $anchor{'pri_provider'} eq 'CELLULAR' ) );
	$anchor{'tad_port_3'}  = 'Optional_3';
	$anchor{'tad_port_3'}  = "$anchor{'mls1_name'}-1" if ( $sitetype =~ /^(K|H)$/ );
	$anchor{'tad_port_3'}  = $anchor{'mls1_name'} if ( $sitetype =~ /^(T|G)$/ );
	$anchor{'tad_port_3'}  = $anchor{'currentstack'} if ( $sitetype =~ /^(D|C|N|F)$/ );
	$anchor{'tad_port_3'}  = $anchor{'cgw2_name'} if ( ( $sitetype =~ /^(U)$/ ) and ($anchor{'isp'} eq 'CELLULAR' and $anchor{'pri_provider'} eq 'CELLULAR' ) );
	$anchor{'tad_port_4'}  = 'Optional_4';
	$anchor{'tad_port_4'}  = $anchor{'vgc1_name'} if ( $anchor{'vgcount'} > 0 );
	$anchor{'tad_port_4'}  = "$anchor{'mls1_name'}-2" if ( $sitetype =~ /^(K|H)$/ );
	$anchor{'tad_port_4'}  = $anchor{'mls2_name'} if ( $sitetype =~ /^(T|G)$/ );
	$anchor{'tad_port_5'}  = 'Optional_5';
	$anchor{'tad_port_5'}  = $anchor{'wlc1_name'}
	if ( $anchor{'wlan'} eq 'Y' and $sitetype !~ /^(D|C|N|F)$/ );
	$anchor{'tad_port_6'}  = '9300_spare';
	$anchor{'tad_port_6'}  = $anchor{'netscout1_name'} if ( $sitetype =~ /^(D|C|N|F)$/ );
	$anchor{'tad_port_7'}  = '9300_spare';
	$anchor{'tad_port_7'}  = 'Optional_7' if ( $sitetype !~ /^(T|G)$/ );
	$anchor{'tad_port_8'}  = 'Optional_8';
	$anchor{'tad_port_8'}  = 'FW1' if ( $anchor{'fw'} eq 'Y' );
	$anchor{'tad_port_9'}  = 'Optional_9';
	$anchor{'tad_port_9'}  = 'FW2' if ( $anchor{'fw'} eq 'Y' );
	$anchor{'tad_port_10'} = 'Optional_10';
	$anchor{'tad_port_10'} = $anchor{'wlc2_name'} if ( ($anchor{'wlan'} eq 'Y') and ($anchor{'wlc_nmbr'} == 2) );
	$anchor{'tad_port_11'} = 'Optional_11';
	$anchor{'tad_port_11'} = $anchor{'vgc1_name'} if ( $anchor{'vgcount'} > 0 );
	$anchor{'tad_port_12'} = 'Optional_12';
	$anchor{'tad_port_12'} = $anchor{'vgc2_name'} if ( $anchor{'vgcount'} == 2 );
	$anchor{'tad_port_13'} = $anchor{'netscout1_name'};
	$anchor{'tad_port_14'} = 'Optional_14';
	$anchor{'tad_port_14'} = $anchor{'sbc1_oracle_name'} if ($anchor{'sbc'} eq 'yes');
	$anchor{'tad_port_15'} = 'Optional_15';
	$anchor{'tad_port_15'} = $anchor{'sbc2_oracle_name'} if ( ($anchor{'sbc'} eq 'yes') and ($sitetype =~ /^(T|G)$/) );
	$anchor{'tad_port_16'} = 'Optional_16';
	$anchor{'tad_int'} = 'Gi1/0/14';
	$anchor{'tad_int'} = 'Gi5/0/14' if ( $sitetype =~ /^(T|G)$/ );
}

sub writeIPSummaryNew {
	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid, my $sitetype, my $stacklimit ) = @_;
	my $mdf_flr = $anchor{'mdf_bldg'} . sprintf( "%02d", $anchor{'mdf_flrnumber'} );
	my $outputFile = $siteid . '-IP-Summary-Chart.xls';

	my $workbook = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile")
	  or die "create XLS file '$SVR_ROOTDIR/$OutputDir/$outputFile' failed: $!";

	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75);
	$ws->set_column( 'A:A', 14.5 );
	$ws->set_column( 'B:B', 16.5 );
	$ws->set_column( 'C:C', 16.5 );
	$ws->set_column( 'D:D', 16.5 );
	$ws->set_column( 'E:E', 32 );
	$ws->set_column( 'F:F', 23 );
	my $fmtBlank = $workbook->add_format( size => 12, bold => 1 );
	my $fmtGray = $workbook->add_format(
										 size      => 9,
										 bold      => 1,
										 bg_color  => 22,
										 text_wrap => 1
	);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtYellow = $workbook->add_format( bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );

	# First three rows are the header
	my $rows   = 2;
	my $height = 12;
	$ws->merge_range( 'A1:H2', "$siteid - IP Summary Chart", $fmtBlank );
	my @header = (
				   'Status',
				   'IDF',
				   'IP Subnet',
				   'IP Address',
				   'Device Name',
				   'Device Description',
				   'DHCP',
				   'DNS',
	);
	$ws->write_row( $rows, 0, \@header, $fmtGray );
	$rows++;

	# CIS values
	my %cistypes = ( 'ASR', 'ASR1001-X', '4321', 'C4321', '4331', 'C4331',
					'4351', 'C4351', '4451', 'C4451', '4461', 'C4461', 'C8200-1N-4T', 'C8200-1N-4T', 'C8300-1N1S', 'C8300-1N1S', 'C8300-2N2S', 'C8300-2N2S', 'C8500-12X4QC', 'C8500-12X4QC',
					'C1111X-8P', 'C1111X-8P', 'C1161X-8P', 'C1161X-8P', 'C1121X-8P', 'C1121X-8P' ),

	my $cistype;
	$cistype = $cistypes{$anchor{'router_type'}};

	# MLS values
	my ( $mlstype, $gwyname );
	$mlstype = 'C9300-48T-A' if ( $sitetype =~ /^(K|H)$/ );
	$mlstype = 'C9407R' if ( $sitetype =~ /^(T|G)$/ );
	$gwyname = "$anchor{'mls1_name'}-vlan-" if ($sitetype =~ /^(K|H)$/ );
	$gwyname = "mls$siteid$mdf_flr" . 'a0x-vlan-' if ( $sitetype =~ /^(T|G)$/ );

	# Decide whether Data scopes or not
	my $vdhcp = 'Yes';
	$vdhcp = 'Yes*' if ( $anchor{'ipt'} eq 'Y' );

	# Get stack name for D, C, N and F sites
	my $stackname   = ( split( /,/, $StackList[0] ) )[0];

	# AT&T CER-PER
	my $wan_subnet;
	if ( $anchor{'pri_wan_ip_cer'} ne '' ) {
		( my $octet1, my $octet2, my $octet3, my $octet4 ) =
		  split( /\./, $anchor{'pri_wan_ip_cer'} );
		$octet4--;
		$wan_subnet = join( '.', $octet1, $octet2, $octet3, $octet4 );
		}

	# Trimming out "anchor" to shorten horizontal length
	my $pri_wan_ip_per = $anchor{'pri_wan_ip_per'};
	my $pri_wan_ip_cer = $anchor{'pri_wan_ip_cer'};
	my $att_upl_int = $anchor{'att_upl_int_dns'};
	my $cis_mls_int = $anchor{'cis_mls_int'};
	my $cis_mls_int2 = $anchor{'cis_mls_int2'};
	my $tad1_ip = $anchor{'tad1_ip'};
	my $loop_subnet = $anchor{'loop_subnet'};
	my $data_subnet_1 = $anchor{'data_subnet_1'};
	my $vgc1_address = $anchor{'vgc1_address'};
	my $vgc2_address = $anchor{'vgc2_address'};
	my $tad1_subnet = $anchor{'tad1_subnet'};
	my $tad1_name = $anchor{'tad1_name'};
	my $vgc1_name = $anchor{'vgc1_name'};
	my $vgc2_name = $anchor{'vgc2_name'};
	my $mls1_name = $anchor{'mls1_name'};
	my $mls2_name = $anchor{'mls2_name'};
	my $cis1_name = $anchor{'cis1_name'};
	my $cis2_name = $anchor{'cis2_name'};
	my $stk_model = $anchor{'stk_model'};
	my $r1_vlan = $anchor{'r1_vlan'};
	my $cgw1_name = $anchor{'cgw1_name'};
	my $cgw2_name = $anchor{'cgw2_name'};


	my ($wlc1_name, $wlc1_mgmt_ip, $wlc2_name, $wlc2_mgmt_ip, $wlan_subnet_i, $wlan_subnet_e, $wlan_gwyip_i, $wemsk,
		$wlan_gwyip_e, $wlan_wlcip_i,$wlan_wlcip_e, $wlan_mls1ip_e, $wlan_mls2ip_e, $wmsk );
	if ( $anchor{'wlan'} eq 'Y' ) {
		$wlc1_name = $anchor{'wlc1_name'};
		$wlc2_name = $anchor{'wlc2_name'};
		$wlc1_mgmt_ip = $anchor{'wlc1_mgmt_ip'};
		$wlc2_mgmt_ip = $anchor{'wlc2_mgmt_ip'};
		$wlan_subnet_i = $anchor{'wlan_subnet_i'};
		$wlan_subnet_e = $anchor{'wlan_subnet_e'};
		$wlan_gwyip_i = $anchor{'wlan_gwyip_i'};
		$wlan_gwyip_e = $anchor{'wlan_gwyip_e'};
		$wlan_wlcip_i = $anchor{'wlan_wlcip_i'};
		$wlan_wlcip_e = $anchor{'wlan_wlcip_e'};
		$wlan_mls1ip_e = $anchor{'wlan_mls1ip_e'};
		$wlan_mls2ip_e = $anchor{'wlan_mls2ip_e'};
		$wmsk = "0/$anchor{'wlan_subnet_i_mask_nexus'}";
		$wmsk = '128/25' if ( $sitetype =~ /^(H|D|C)$/ );
		$wemsk = $anchor{'wlan_subnet_e_mask_nexus'};
	}

	# Start writing output
	my @row = ( 'New', 'mdf', '', '', '', '', 'No', 'Yes' );
	my @ipsum = (
		#-[0]----------------------[1]-----------------[2]------------------------[3]-------------------[4]------------------#
		[ $tad1_subnet, 		   $tad1_ip,  		   $tad1_name, 				  'SLC8000', 			'TGKHDN' ],#ip ==  0
		[ "$loop_subnet.1/32",     "$loop_subnet.1",   $cis1_name,				  $cistype , 			'TGKHDCNF' ],#ip ==  1
		[ "$loop_subnet.2/32",     "$loop_subnet.2",   $cis2_name, 				  $cistype ,  			'TGKHDN'   ],#ip ==  2
		[ "$loop_subnet.7/32",     "$loop_subnet.7",   $mls1_name, 				  $mlstype, 			'KH'  	   ],#ip ==  3
		[ "$loop_subnet.14/32",    "$loop_subnet.14",  $mls1_name, 				  $mlstype , 			'TG' 	   ],#ip ==  4
		[ "$loop_subnet.15/32",    "$loop_subnet.15",  $mls2_name, 				  $mlstype  ,  			'TG'	   ],#ip ==  5
		[ "$data_subnet_1.0/24",   $vgc1_address, 	   $vgc1_name, 				  'VG204XM'	, 			'DC'	   ],#ip ==  6
		[ "$loop_subnet.128/25",   $vgc1_address, 	   $vgc1_name, 				  'VG204XM', 			'TNF'	   ],#ip ==  7
		[ "$loop_subnet.64/26",    $vgc1_address, 	   $vgc1_name, 				  'VG204XM', 			'G'		   ],#ip ==  8
		[ "$loop_subnet.32/27",    $vgc1_address, 	   $vgc1_name, 				  'VG204XM'	,			'KH'	   ],#ip ==  9
		[ "$data_subnet_1.0/24",   $vgc2_address, 	   $vgc2_name, 				  'VG204XM', 			'DC'	   ],#ip == 10
		[ "$loop_subnet.128/25",   $vgc2_address, 	   $vgc2_name, 				  'VG204XM', 			'TNF'	   ],#ip == 11
		[ "$loop_subnet.64/26",    $vgc2_address, 	   $vgc2_name, 				  'VG204XM', 			'G'		   ],#ip == 12
		[ "$loop_subnet.32/27",    $vgc2_address, 	   $vgc2_name, 				  'VG204XM'	,			'KH'	   ],#ip == 13
		[ "$loop_subnet.128/25",   $wlc1_mgmt_ip, 	   $wlc1_name, 				  'WLC5500', 			'T'		   ],#ip == 14
		[ "$loop_subnet.64/26",    $wlc1_mgmt_ip, 	   $wlc1_name, 				  'WLC5500', 			'G'		   ],#ip == 15
		[ "$loop_subnet.32/27",    $wlc1_mgmt_ip,	   $wlc1_name, 				  'WLC5500', 			'KH'	   ],#ip == 16
		[ "$wan_subnet/30",		   $pri_wan_ip_cer,    "$cis1_name-$att_upl_int$r1_vlan", 'WAN Circuit', 		'TGKHDCNF' ],#ip == 17
		[ "$wan_subnet/30",		   $pri_wan_ip_per,	   "per-$cis1_name",		  'WAN Circuit', 		'TGKHDCNF' ],#ip == 18
		[ "$loop_subnet.16/29",    "$loop_subnet.17",  "$cis1_name-g0-0-1-100",   $cistype, 			'KHDCNF'   ],#ip == 19
		[ "$loop_subnet.16/29",	   "$loop_subnet.18",  "$cis2_name-g0-0-1-100",   $cistype , 			'KHDN' 	   ],#ip == 20
		[ "$loop_subnet.16/29",    "$loop_subnet.19",  "$stackname-vlan-100" ,	  $stk_model, 			'DCNF'	   ],#ip == 21
		[ "$loop_subnet.16/29",    "$loop_subnet.19",  "$mls1_name-vlan-100",     $mlstype, 			'KH'	   ],#ip == 22
		[ "$loop_subnet.32/30",    "$loop_subnet.33",  "$cis1_name-$cis_mls_int", 	  'cis1 to mls1', 		'TG' 	   ],#ip == 23
		[ "$loop_subnet.32/30",    "$loop_subnet.34",  "$mls1_name-g5-0-1", 	  'mls1 to cis1', 		'TG'   	   ],#ip == 24
		[ "$loop_subnet.36/30",    "$loop_subnet.37",  "$cis1_name-$cis_mls_int2", 	  'cis1 to mls2', 		'TG'       ],#ip == 25
		[ "$loop_subnet.36/30",    "$loop_subnet.38",  "$mls2_name-g5-0-2", 	  'mls2 to cis1', 		'TG' 	   ],#ip == 26
		[ "$loop_subnet.40/30",    "$loop_subnet.41",  "$cis2_name-$cis_mls_int2", 	  'cis2 to mls1', 		'TG' 	   ],#ip == 27
		[ "$loop_subnet.40/30",    "$loop_subnet.42",  "$mls1_name-g5-0-2", 	  'mls1 to cis2', 		'TG'	   ],#ip == 28
		[ "$loop_subnet.44/30",    "$loop_subnet.45",  "$cis2_name-$cis_mls_int", 	  'cis2 to mls2', 		'TG'	   ],#ip == 29
		[ "$loop_subnet.44/30",    "$loop_subnet.46",  "$mls2_name-g5-0-1",   	  'mls2 to cis2', 		'TG'	   ],#ip == 30
		[ "$loop_subnet.48/30",    "$loop_subnet.49",  "$mls1_name-vlan-100",	  'mls1 to mls2', 		'TG'	   ],#ip == 31
		[ "$loop_subnet.48/30",    "$loop_subnet.50",  "$mls2_name-vlan-100", 	  'mls2 to mls1',		'TG'	   ],#ip == 32
		[ $tad1_subnet, 		   "$loop_subnet.58",  "$mls1_name-vlan-901", 	  'mls1 to tad', 		'TG'	   ],#ip == 33
		[ "$loop_subnet.8/29",     "$loop_subnet.9",   $gwyname . '99-gwy',	      'Firewall Vlan Gwy',  'KH'       ],#ip == 34
		[ "$loop_subnet.16/28",    "$loop_subnet.17",  $gwyname . '99-gwy', 	  'Firewall Vlan Gwy',  'TG'       ],#ip == 35
		[ '', 					   "$loop_subnet.18",  "$mls1_name-vlan-99",	  'Firewall Vlan',		'TG'	   ],#ip == 36
		[ '',					   "$loop_subnet.19",  "$mls2_name-vlan-99",	  'Firewall Vlan', 		'TG'	   ],#ip == 37
		[ "$loop_subnet.32/27",    "$loop_subnet.33",  $gwyname . '101-gwy', 	  'Server Vlan Gwy',	'HK'       ],#ip == 38
		[ "$loop_subnet.64/26",    "$loop_subnet.65",  $gwyname . '101-gwy', 	  'Server Vlan Gwy',	'G'        ],#ip == 39
		[ '', 					   "$loop_subnet.66",  "$mls1_name-vlan-101",	  'Server Vlan', 		'G'	 	   ],#ip == 40
		[ '',					   "$loop_subnet.67",  "$mls2_name-vlan-101",	  'Server Vlan', 		'G'	 	   ],#ip == 41
		[ "$loop_subnet.128/25",   "$loop_subnet.129", $gwyname . '101-gwy', 	  'Server Vlan Gwy',	'T'        ],#ip == 42
		[ '', 					   "$loop_subnet.130", "$mls1_name-vlan-101",	  'Server Vlan',		'T'	 	   ],#ip == 43
		[ '',					   "$loop_subnet.131", "$mls2_name-vlan-101",	  'Server Vlan', 		'T'	 	   ],#ip == 44
		[ "$loop_subnet.$wmsk",	   "$loop_subnet.129", "$stackname-vlan-110-gwy", 'WLAN User Vlan Gwy', 'DC'	   ],#ip == 45
		[ "$wlan_subnet_i.$wmsk",  $wlan_gwyip_i,	   $gwyname . '110-gwy', 	  'WLAN User Vlan Gwy', 'TGKH'	   ],#ip == 46
		[ '',  					   "$wlan_subnet_i.3", "$mls1_name-vlan-110", 	  'WLAN User Vlan', 	'TG' 	   ],#ip == 47
		[ '',  					   "$wlan_subnet_i.2", "$mls2_name-vlan-110", 	  'WLAN User Vlan',     'TG'	   ],#ip == 48
		[ '',  					   $wlan_wlcip_i,	   "$wlc1_name-vlan-110", 	  'WLC5500', 			'TGKH' 	   ],#ip == 49
		[ "$wlan_subnet_e/$wemsk", $wlan_gwyip_e,	   $gwyname . '119-gwy', 	  'WLAN EAC Vlan Gwy',  'TGKH'     ],#ip == 50
		[ '',  					   $wlan_mls1ip_e,     "$mls1_name-vlan-119", 	  'WLAN EAC Vlan', 		'TG'       ],#ip == 51
		[ '',  					   $wlan_mls2ip_e,	   "$mls2_name-vlan-119", 	  'WLAN EAC Vlan' , 	'TG'	   ],#ip == 52
		[ '',   				   $wlan_wlcip_e,	   "$wlc1_name-vlan-119", 	  'WLC5500', 		    'TGKH' 	   ], #ip == 53
		[ "$anchor{'utgw'}/31", 		   $anchor{'utad'},  		   $tad1_name, 				  'EMG8500', 			'U' ],#ip ==  54
		[ "$anchor{'ucis'}/32",     $anchor{'ucis'},   $cis1_name,				  $cistype , 			'U' ],#ip ==  55
		[ "$anchor{'ucgw1'}/32",     $anchor{'ucgw1'},   $cgw1_name,				  'CG522-E' , 			'U' ],#ip ==  56
		[ $tad1_subnet, 		   $tad1_ip,  		   $tad1_name, 				  'EMG8500', 			'CF' ]#ip ==  57

	);
	for ( my $ip = 0; $ip <= $#ipsum; $ip++ ) { # $ip will be used as parent array's index
		$row[2] = $ipsum[$ip][0];
		$row[3] = $ipsum[$ip][1];
		$row[4] = $ipsum[$ip][2];
		$row[5] = $ipsum[$ip][3];
		$row[1] = 'mdf';
		$row[1] = 'Vlan101/mdf' if ( $row[4] =~ /vlan-101/ );
		$row[1] = 'Vlan100/mdf' if ( $row[4] =~ /vlan-100/ );
		$row[1] = 'Vlan901/mdf' if ( $row[4] =~ /vlan-901/ );
		$row[1] = 'Vlan99/mdf' if ( $row[4] =~ /vlan-99/ );
		$row[1] = 'Vlan110/mdf' if ( $row[4] =~ /vlan-110/ );
		$row[1] = 'Vlan119/mdf' if ( $row[4] =~ /vlan-119/ );
		$row[6] = 'No';
		$row[6] = 'Yes' if ( $row[4] =~ /vlan-110-gwy/ );
		$row[6] = 'Yes**' if ( $row[4] =~ /vlan-119-gwy/ );
		my $fmtColor = $fmtPurple;
		$fmtColor = $fmtBlue if ( $row[4] =~ /-/ );
		$fmtColor = $fmtYellow if ( $row[4] =~ /(-110|-119)/ );

		if ( $ipsum[$ip][4] =~ /$sitetype/ and $row[4] =~ /(tad|cis|mls|stk|vgc|wlc)/ ) {
			$rows = writeRow( $ws, $rows, \@row, $fmtColor )
			unless ( ( $row[4] =~ /vlan-99/ and $anchor{'fw'} ne 'Y' )
				or ( $row[5] =~ /WLAN/ and $anchor{'wlan'} ne 'Y' )
	      or ( $row[5] eq 'WLC5500' )); # Preempt creating entry for 5500s if wlc is 9800);
		}

		# WLC Device description
		if ( $anchor{'wlc_model'} eq '9800' and $ipsum[$ip][3] eq 'WLC5500' and $ipsum[$ip][4] =~ /$sitetype/){
			$row[5] = 'WLC9800';
			$rows = writeRow( $ws, $rows, \@row, $fmtColor );
		}elsif( $anchor{'wlc_model'} ne '9800' and $ipsum[$ip][3] eq 'WLC5500' and $ipsum[$ip][4] =~ /$sitetype/){
			$rows = writeRow( $ws, $rows, \@row, $fmtColor );
		}

		#Selection for 2nd WLC - 9800 (T,G,K,H)
		if ( $ip == 14 and $anchor{'wlc_nmbr'} == 2 and ($sitetype =~ /^(T)$/) ){
			$row[2] = "$loop_subnet.128/25";
			$row[3] = $wlc2_mgmt_ip;
			$row[4] = $wlc2_name;
			$row[5] = 'WLC9800';
			$rows   = writeRow( $ws, $rows, \@row, $fmtColor );
		}	if ( $ip == 15 and $anchor{'wlc_nmbr'} == 2 and ($sitetype =~ /^(G)$/) ){
			$row[2] = "$loop_subnet.64/26";
			$row[3] = $wlc2_mgmt_ip;
			$row[4] = $wlc2_name;
			$row[5] = 'WLC9800';
			$rows   = writeRow( $ws, $rows, \@row, $fmtColor );
		}	if ( $ip == 16 and $anchor{'wlc_nmbr'} == 2 and ($sitetype =~ /^(K|H)$/) ){
			$row[2] = "$loop_subnet.32/27";
			$row[3] = $wlc2_mgmt_ip;
			$row[4] = $wlc2_name;
			$row[5] = 'WLC9800';
			$rows   = writeRow( $ws, $rows, \@row, $fmtColor );
		}

		if ( $ip == 33  and $anchor{'SDWAN'} eq 'Y' ) {
		$anchor{'tloc_int'} =~ s/\//-/g;
		my $tloc_int_dns = lc($anchor{'tloc_int'});
		$anchor{'tloc_int2'} =~ s/\//-/g;
		my $tloc_int2_dns = lc($anchor{'tloc_int2'});
		$anchor{'tloc_int3'} =~ s/\//-/g;
		my $tloc_int3_dns = lc($anchor{'tloc_int3'});
		$anchor{'tloc_int4'} =~ s/\//-/g;
		my $tloc_int4_dns = lc($anchor{'tloc_int4'});

		# SDWAN TLOC Extensions
		if( $anchor{'tloc'} eq 'yes' ){
			if ( $sitetype !~ /^(C|F|U)$/ ) {
				if ( $anchor{'transport'} == 2){
				# Router 1
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.24/30";
				$row[2] = "$loop_subnet.52/30" if ( $sitetype =~ /^(T|G)$/ );
				$row[3] = "$loop_subnet.25";
				$row[3] = "$loop_subnet.53" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis1_name-$tloc_int_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.26";
				$row[3] = "$loop_subnet.54" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis2_name-$tloc_int_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2
				$row[2] = "$loop_subnet.28/30";
				$row[2] = "$loop_subnet.60/30" if ( $sitetype =~ /^(T|G)$/ );
				$row[3] = "$loop_subnet.29";
				$row[3] = "$loop_subnet.61" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis1_name-$tloc_int2_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.30";
				$row[3] = "$loop_subnet.62" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis2_name-$tloc_int2_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
				if ( $anchor{'transport'} == 3){
				# Router 1
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.24/30";
				$row[2] = "$loop_subnet.52/30" if ( $sitetype =~ /^(T|G)$/ );
				$row[3] = "$loop_subnet.25";
				$row[3] = "$loop_subnet.53" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis1_name-$tloc_int_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.26";
				$row[3] = "$loop_subnet.54" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis2_name-$tloc_int_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2
				$row[2] = "$loop_subnet.28/30";
				$row[2] = "$loop_subnet.60/30" if ( $sitetype =~ /^(T|G)$/ );
				$row[3] = "$loop_subnet.29";
				$row[3] = "$loop_subnet.61" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis1_name-$tloc_int2_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.30";
				$row[3] = "$loop_subnet.62" if ( $sitetype =~ /^(T|G)$/ );
				$row[4] = "$cis2_name-$tloc_int2_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 1 - MPLS,Private1,Private2 TLOC-interface
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.96/30";
				$row[2] = "$loop_subnet.192/30" if ( $sitetype =~ /^(G)$/ );
				$row[2] = "$loop_subnet.64/30" if ( $sitetype =~ /^(T)$/ );
				$row[3] = "$loop_subnet.97";
				$row[3] = "$loop_subnet.193" if ( $sitetype =~ /^(G)$/ );
				$row[3] = "$loop_subnet.65" if ( $sitetype =~ /^(T)$/ );
				$row[4] = "$cis1_name-$tloc_int3_dns";
				$row[5] = 'cis1 TLOC extension';
				$anchor{'r1_mpls_tloc_sub'} = $row[2];
				$anchor{'r1_mpls_tloc_ip'} = $row[3];
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.98";
				$row[3] = "$loop_subnet.194" if ( $sitetype =~ /^(G)$/ );
				$row[3] = "$loop_subnet.66" if ( $sitetype =~ /^(T)$/ );
				$row[4] = "$cis2_name-$tloc_int3_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$anchor{'r2_mpls_tloc_ip'} = $row[3];
				}
				if (( $anchor{'transport'} == 4) and ( $sitetype =~ /^(T|G)$/ )) {
				# Router 1
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.52/30";
				$row[3] = "$loop_subnet.53";
				$row[4] = "$cis1_name-$tloc_int_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.54";
				$row[4] = "$cis2_name-$tloc_int_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2
				$row[2] = "$loop_subnet.60/30";
				$row[3] = "$loop_subnet.61";
				$row[4] = "$cis1_name-$tloc_int2_dns";
				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.62";
				$row[4] = "$cis2_name-$tloc_int2_dns";
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 1 - MPLS,Private1,Private2 TLOC-interface
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.192/30" if ( $sitetype =~ /^(G)$/ );
				$row[2] = "$loop_subnet.64/30" if ( $sitetype =~ /^(T)$/ );
				$row[3] = "$loop_subnet.193" if ( $sitetype =~ /^(G)$/ );
				$row[3] = "$loop_subnet.65" if ( $sitetype =~ /^(T)$/ );
				$row[4] = "$cis1_name-$tloc_int3_dns";
				$row[5] = 'cis1 TLOC extension';
				$anchor{'r1_mpls_tloc_sub'} = $row[2];
				$anchor{'r1_mpls_tloc_ip'} = $row[3];
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.194" if ( $sitetype =~ /^(G)$/ );
				$row[3] = "$loop_subnet.66" if ( $sitetype =~ /^(T)$/ );
				$row[4] = "$cis2_name-$tloc_int3_dns";
				$row[5] = 'cis2 TLOC extension';
				$anchor{'r2_mpls_tloc_ip'} = $row[3];
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# Router 2 - MPLS,Private1,Private2 TLOC-interface
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.196/30" if ( $sitetype =~ /^(G)$/ );
				$row[2] = "$loop_subnet.68/30" if ( $sitetype =~ /^(T)$/ );
				$row[3] = "$loop_subnet.197" if ( $sitetype =~ /^(G)$/ );
				$row[3] = "$loop_subnet.69" if ( $sitetype =~ /^(T)$/ );
				$row[4] = "$cis1_name-$tloc_int4_dns";
				$row[5] = 'cis1 TLOC extension';
				$anchor{'r2_mpls2_tloc_sub'} = $row[2];
				$anchor{'r2_mpls2_tloc_ip'} = $row[3];
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[3] = "$loop_subnet.198" if ( $sitetype =~ /^(G)$/ );
				$row[3] = "$loop_subnet.70" if ( $sitetype =~ /^(T)$/ );
				$row[4] = "$cis2_name-$tloc_int4_dns";
				$row[5] = 'cis2 TLOC extension';
				$anchor{'r1_mpls2_tloc_ip'} = $row[3];
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
			}
		}
		else{
		# Values change in accordance to site type, transport count and router model.
			if ( $sitetype !~ /^(C|F|U)$/ ) {
				# Router 1
				$row[1] = 'mdf';
				$row[2] = "$loop_subnet.28/30";
				$row[2] = "$loop_subnet.52/30" if ( $sitetype =~ /^(T|G)$/ );
				$row[3] = "$loop_subnet.29";
				$row[3] = "$loop_subnet.53" if ( $sitetype =~ /^(T|G)$/ );

				$row[4] = "$cis1_name-$tloc_int_dns";

				$row[5] = 'cis1 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				# Router 2
				$row[3] = "$loop_subnet.30";
				$row[3] = "$loop_subnet.54" if ( $sitetype =~ /^(T|G)$/ );
				if ( $anchor{'transport'} == 2){
					$row[4] = "$cis2_name-$tloc_int_dns";
				}
				else{
					$row[4] = "$cis2_name-$tloc_int2_dns";
				}
				$row[5] = 'cis2 TLOC extension';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				# 2 transports only
				if ( $anchor{'transport'} == 2 and $sitetype =~ /^(T|G|H|D|N)$/ ) {
					if ($sitetype =~ /^(H|D|N)$/){
						$row[2] = "$loop_subnet.24/30";
						$row[3] = "$loop_subnet.25";
						#$row[4] = "$cis1_name-g0-0-2-20";
						$row[4] = "$cis1_name-$tloc_int2_dns";
						$row[5] = 'cis1 TLOC extension';
						# prtout("IP summary Row4: $row[4]");
						$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
						$row[3] = "$loop_subnet.26";
						#$row[4] = "$cis2_name-g0-0-2-20";
						$row[4] = "$cis2_name-$tloc_int2_dns";
						$row[5] = 'cis2 TLOC extension';
						# prtout("IP summary Row4: $row[4]");
						$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
					}
					if ($sitetype =~ /^(T|G)$/){
						$row[2] = "$loop_subnet.60/30";
						$row[3] = "$loop_subnet.61";
						#$row[4] = "$cis1_name-g0-0-2-20";
						$row[4] = "$cis1_name-$tloc_int2_dns";
						$row[5] = 'cis1 TLOC extension';
						# prtout("IP summary Row4: $row[4]");
						$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
						$row[3] = "$loop_subnet.62";
						#$row[4] = "$cis2_name-g0-0-2-20";
						$row[4] = "$cis2_name-$tloc_int2_dns";
						$row[5] = 'cis2 TLOC extension';
						# prtout("IP summary Row4: $row[4]");
						$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
					}
				}
			}
			}
		}
		#SBC settings for VLANs 397 and 398
		if ( ($ip == 44) and ($anchor{'sbc'} =~ 'yes') and $sitetype =~ /^(T|G|K|H)$/ ){
				#Vlan397
				$row[1] = 'Vlan397/mdf';
				$row[2] = "$anchor{'sbc_sub'}.128/26";
				$row[3] = "$anchor{'sbc_397gw'}";
				$row[4] = $gwyname.'397-gwy';
				$row[5] = 'SBC Vlan397 Gwy';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				if ($sitetype =~ /^(T|G)$/){
				$row[1] = 'Vlan397/mdf';
				$row[2] = '';
				$row[3] = "$anchor{'sbc_397_ip1'}";
				$row[4] = $mls1_name.'-vlan-397';
				$row[5] = 'SBC Vlan397';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[1] = 'Vlan397/mdf';
				$row[2] = '';
				$row[3] = "$anchor{'sbc_397_ip2'}";
				$row[4] = $mls2_name.'-vlan-397';
				$row[5] = 'SBC Vlan397';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
				#Vlan398
				$row[1] = 'Vlan398/mdf';
				$row[2] = "$anchor{'sbc_sub'}.224/28";
				$row[2] = "$anchor{'sbc_sub'}.112/28" if ( $sitetype =~ /^(H)$/ );
				$row[3] = "$anchor{'sbc_398gw'}";
				$row[4] = $gwyname.'398-gwy';
				$row[5] = 'SBC Vlan398 Gwy';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				if ($sitetype =~ /^(T|G)$/){
				$row[1] = 'Vlan398/mdf';
				$row[2] = '';
				$row[3] = "$anchor{'sbc_398_ip1'}";
				$row[4] = $mls1_name.'-vlan-398';
				$row[5] = 'SBC Vlan398';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				$row[1] = 'Vlan398/mdf';
				$row[2] = '';
				$row[3] = "$anchor{'sbc_398_ip2'}";
				$row[4] = $mls2_name.'-vlan-398';
				$row[5] = 'SBC Vlan398';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
		}
		# IDFs
		if ( $ip == 44 and $sitetype !~ /^(U)$/ ) {
			for ( my $ct = 0; $ct <= $#StackList; $ct++ ) {
				my $stack         = $ct + 1;
				my $stackname   = ( split( /,/, $StackList[$ct] ) )[0];
				my $stackvlan	  = $anchor{"idf$stack" . '_Data_vlan_1'};
				my $stackevlan	  = $anchor{"idf$stack" . '_EAC_vlan_1'};
				my $stacksubnet   = $anchor{"idf$stack" . '_ds_1'};
				my $stackesubnet  = $anchor{"idf$stack" . '_es_1'};
				my $stacklastoc1  = $anchor{"idf$stack" . '_dlo_1'};
				my $stacklastoc2  = $anchor{"idf$stack" . '_dlo_2'};
				my $stackelastoc1 = $anchor{"idf$stack" . '_elo_1'};
				my $stackelastoc2 = $anchor{"idf$stack" . '_elo_2'};
				my $stackegwy 	  = $anchor{"idf$stack" . '_ego'};
				my $emask		  = $anchor{"idf$stack" . '_eno'} . '/27';
				my @vlanset       = ( $stackvlan,   $stackevlan, '201', '401' );
				my @subnetset     = ( $stacksubnet, $stackesubnet, $data_subnet_1, $loop_subnet );
				my @maskset       = ( '0/24', $emask, '0/24', '32/27' );
				my @gwyset        = ( '1', $stackegwy, '1', '33' );
				my @nameset       = ( '-gwy', '-gwy-eac' );
				my @lastoc1set    = ( $stacklastoc1, $stackelastoc1 );
				my @lastoc2set    = ( $stacklastoc2, $stackelastoc2 );
				my @colorset      = ( $fmtGreen, $fmtTeal );
				my @dhcpset		  = ( $vdhcp, 'Yes' );

				for ( my $ct2 = 0; $ct2 <= 1; $ct2++ ) {
					# mls1 for K and H sites
					# HSRP for T and G sites
					$row[1] = "$vlanset[$ct2]/idf";
					$row[2] = "$subnetset[$ct2].$maskset[$ct2]";
					$row[3] = "$subnetset[$ct2].$gwyset[$ct2]";
					$row[4] = "$gwyname$vlanset[$ct2]$nameset[$ct2]";
					$row[5] = "$mlstype";
					# stk for D, C, N  and F sites
					if ( $sitetype =~ /^(D|C|N|F)$/ ) {
						$row[1] = "$vlanset[$ct2 + 2]/idf";
						$row[2] = "$subnetset[$ct2 + 2].$maskset[$ct2 + 2]";
						$row[3] = "$subnetset[$ct2 + 2].$gwyset[$ct2 + 2]";
						$row[2] = "$loop_subnet.128/25" if ( $ct2 == 0 and $sitetype =~ /^(N|F)$/ );
						$row[3] = "$loop_subnet.129"    if ( $ct2 == 0 and $sitetype =~ /^(N|F)$/ );
						$row[4] = "$stackname-vlan-$vlanset[$ct2 + 2]-gwy";
						$row[5] = $stk_model;
					}
					$row[6] = $dhcpset[$ct2];
					$rows   = writeRow( $ws, $rows, \@row, $colorset[$ct2] );
					#Add mgmt IP to DCNF stacks as advised by Bob Benson
					if ( $sitetype =~ /^(D|C|N|F)$/ and $ct2 == 0 ) {
						$row[1] = "$vlanset[$ct2 + 2]/idf";
						$row[2] = "$subnetset[$ct2 + 2].$maskset[$ct2 + 2]";
						$row[3] = "$subnetset[$ct2 + 2].4";
						$row[2] = "$loop_subnet.128/25" if ( $sitetype =~ /^(N|F)$/ );
						$row[3] = "$loop_subnet.132"    if ( $sitetype =~ /^(N|F)$/ );
						$row[4] = $stackname;
						$row[5] = $stk_model;
						$row[6] = $dhcpset[$ct2];
						$rows   = writeRow( $ws, $rows, \@row, $colorset[$ct2] );
					}
					if ( $sitetype =~ /^(T|G)$/ ) {
						# mls1 for T and G sites
						$row[2] = '';
						$row[3] = "$subnetset[$ct2].$lastoc1set[$ct2]";
						$row[4] = "$mls1_name-vlan-$vlanset[$ct2]";
						$row[5] = $mlstype;
						$row[6] = 'No';
						$rows   = writeRow( $ws, $rows, \@row, $colorset[$ct2] );
						# mls2 for T and G sites
						$row[3] = "$subnetset[$ct2].$lastoc2set[$ct2]";
						$row[4] = "$mls2_name-vlan-$vlanset[$ct2]";
						$row[5] = $mlstype;
						$rows   = writeRow( $ws, $rows, \@row, $colorset[$ct2] );
					}
					# stk for T, G, K, and H sites
					if ( $ct2 == 0 and $sitetype =~ /^(T|G|K|H)$/ ) {
						$row[2] = '';
						$row[3] = "$stacksubnet.4";
						$row[4] = $stackname;
						$row[5] = $stk_model;
						$row[6] = 'No';
						$rows   = writeRow( $ws, $rows, \@row, $fmtGreen );
					}
				}
			}
		}

		if ( ($ip == 56) and ($sitetype =~ /^(U)$/) ){

			if ( $anchor{'isp'} eq 'CELLULAR' or $anchor{'pri_provider'} eq 'CELLULAR' ) { # One CG
				$row[2] = "$anchor{'ucgw1'}/32";
				$row[3] = $anchor{'ucgw1'};
				$row[4] = "$cgw1_name";
				$row[5] = 'CG522-E';
				$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );

				if ( $anchor{'isp'} eq 'CELLULAR' and $anchor{'pri_provider'} eq 'CELLULAR' ) { # Two CGs
					$row[2] = "$anchor{'ucgw2'}/32";
					$row[3] = $anchor{'ucgw2'};
					$row[4] = "$cgw2_name";
					$row[5] = 'CG522-E';
					$rows   = writeRow( $ws, $rows, \@row, $fmtBlue );
				}
			}
			$row[1] = 'mdf/401';
			$row[2] = "$anchor{'uena'}/29";
			$row[3] = $anchor{'uegw'};
			$row[4] = "$cis1_name-vlan-401-gwy";
			$row[5] = $cistype;
			$row[6] = 'Yes';
			$rows   = writeRow( $ws, $rows, \@row, $fmtTeal );

			$row[1] = 'mdf/201';
			$row[2] = "$anchor{'udna'}/28";
			$row[3] = $anchor{'udgw'};
			$row[4] = "$cis1_name-vlan-201-gwy";
			$row[5] = $cistype;
			$row[6] = 'Yes';
			$rows   = writeRow( $ws, $rows, \@row, $fmtGreen );
		}
	}

	$rows++;    # leave a blank line before the notes and legend
	@row = ( 'Key', 'if applicable---->', '* - Requires option 150 TFTP Servers' );
	$rows = writeRow( $ws, $rows, \@row );

	# Formats are different for the cells in this row, so write them directly instead of passing to the sub
	$ws->write( $rows, 0, 'Loopback/Mgmt', $fmtPurple );
	$ws->write( $rows, 2, '* - Please add option 150 for the voice DHCP scope per IPT Engineering Team Standards' );
	$rows++;
	$ws->write( $rows, 0, 'Router/MLS', $fmtBlue );
	$ws->write( $rows, 2, '** - Wireless LAN DHCP scopes should be configured with a 2-hour lease time' );
	$rows++;
	$ws->write( $rows, 0, 'User/Stacks', $fmtGreen );
	$ws->write( $rows, 2, 'For a /27, exclude the 1st 7 addresses from DHCP scope.' );
	$rows++;
	$ws->write( $rows, 0, 'EAC', $fmtTeal );
	$ws->write( $rows, 2, 'For a /26, exclude the 1st 10 addresses from DHCP scope.' );
	$rows++;
	$ws->write( $rows, 0, 'Wireless', $fmtYellow );
	$ws->write( $rows, 2, 'For a /25, exclude the 1st 20 addresses from DHCP scope.' );
	$rows++;
	$ws->write( $rows, 2, 'For all others, exclude the 1st 32 addresses from DHCP scope.' );

	return $outputFile;
}

sub writeCISummaryNew {
	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid, my $sitetype, my $stacklimit ) = @_;
	my $outputFile = $siteid . '-CI-Devices.xls';

	my $workbook = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile")
	  or die "create XLS file '$SVR_ROOTDIR/$OutputDir/$outputFile' failed: $!";

	prtout("Writing CI Summary XLS");
	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75); $ws->set_column( 'A:A', 14.5 );
	$ws->set_column( 'B:B', 30 ); $ws->set_column( 'C:C', 17 );
	$ws->set_column( 'D:D', 18 ); $ws->set_column( 'E:E', 25 );
	$ws->set_column( 'F:F', 16.5 );	$ws->set_column( 'G:G', 16.5 );
	$ws->set_column( 'H:H', 18 ); $ws->set_column( 'I:I', 23 );
	$ws->set_column( 'J:J', 13 ); $ws->set_column( 'K:K', 30 );
	$ws->set_column( 'L:L', 12 ); $ws->set_column( 'M:M', 12 );
	my $fmtBlank = $workbook->add_format( size => 12, bold => 1 );
	my $fmtGray = $workbook->add_format(
								size => 9, bold => 1,
								bg_color => 22, text_wrap => 1
							);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtYellow = $workbook->add_format( bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );

	# First three rows are the header
	my $rows   = 2;
	my $height = 12;
	$ws->merge_range( 'A1:O2', "$siteid - CI Summary Chart", $fmtBlank );
	my @header = (
		'If Update to existing CI, provide CI ID',
		'SYSTEM ID/FQDN (OVO ID)', 'FORMAL NAME',
		'COMMON NAME', 'CI TYPE', 'ENVIRONMENT',
		'SERIAL NUMBER', 'MODEL', 'BRAND', 'LOCATION',
		'SUPPORTING WORKGROUP', 'REMARKS', 'IP ADDRESS',
		'IP ADDRESS (additional'
	);
	$rows = writeRow( $ws, $rows, \@header, $fmtGray );

	# CIS values
	my %cistypes = (	'ASR', 'ASR1001-X', '4321', 'C4321', '4331', 'C4331',
					'4351', 'C4351', '4451', 'C4451', '4461', 'C4461', 'C8200-1N-4T', 'C8200-1N-4T', 'C8300-1N1S', 'C8300-1N1S', 'C8300-2N2S', 'C8300-2N2S', 'C8500-12X4QC', 'C8500-12X4QC',
					'C1111X-8P', 'C1111X-8P', 'C1161X-8P', 'C1161X-8P', 'C1121X-8P', 'C1121X-8P' ),

	my $cistype;
	$cistype = $cistypes{$anchor{'router_type'}};

	# MLS values
	my $mlstype = '!';
	$mlstype = 'C9300-48T-A' if ( $sitetype =~ /^(K|H)$/ );
	$mlstype = 'C9407R' if ( $sitetype =~ /^(T|G)$/ );

	#WLC Values
	my $wlctype;
	if ( $anchor{'wlc_model'} eq '9800'){
		$wlctype = 'WLC9800';
	}elsif ( $anchor{'wlc_model'} ne '9800'){
		$wlctype = 'WLC5500';
	}

	# Supporting workgroup is always the same
	my $workgroup = 'NHS_SLO_Remote'; # Per Sonja 04/30/2020
	# CI values
	my @civals = (
		#-[0]-------------------[1]--------[2]---------------------------[3]---------[4]-------------------#
		[ $anchor{'tad1_name'}, 'SLC8000', $anchor{'tad1_ip'}, 		     $fmtPurple, 'TGKHDN' ],#$ci ==  0
		[ $anchor{'vgc1_name'}, 'VG204',   $anchor{'vgc1_address'}, 	 $fmtPurple, 'TGKHDCNF' ],#$ci ==  1
		[ $anchor{'vgc2_name'}, 'VG204',   $anchor{'vgc2_address'}, 	 $fmtPurple, 'TGKHDCNF' ],#$ci ==  2
		[ $anchor{'cis1_name'}, $cistype,  "$anchor{'loop_subnet'}.1",   $fmtPurple, 'TGKHDCNF' ],#$ci ==  3
		[ $anchor{'cis2_name'}, $cistype,  "$anchor{'loop_subnet'}.2",   $fmtPurple, 'TGKHDN' 	],#$ci ==  4
		[ $anchor{'mls1_name'}, $mlstype,  "$anchor{'loop_subnet'}.7" ,  $fmtPurple, 'KH' 		],#$ci ==  5
		[ $anchor{'mls1_name'}, $mlstype,  "$anchor{'loop_subnet'}.14",  $fmtPurple, 'TG' 		],#$ci ==  6
		[ $anchor{'mls2_name'}, $mlstype,  "$anchor{'loop_subnet'}.15",  $fmtPurple, 'TG' 		],#$ci ==  7
		[ $anchor{'wlc1_name'}, $wlctype, "$anchor{'loop_subnet'}.38",  $fmtYellow, 'KH' 		],#$ci ==  8
		[ $anchor{'wlc1_name'}, $wlctype, "$anchor{'loop_subnet'}.70",  $fmtYellow, 'G' 		],#$ci ==  9
		[ $anchor{'wlc1_name'}, $wlctype, "$anchor{'loop_subnet'}.134", $fmtYellow, 'T' 		], #$ci == 10
		[ $anchor{'tad1_name'}, 'EMG8500', $anchor{'utad'}, 		     $fmtPurple, 'U' ],#$ci ==  11
		[ $anchor{'cis1_name'}, $cistype,  $anchor{'ucis'},   			$fmtPurple, 'U' ],#$ci ==  12
		[ $anchor{'cgw1_name'}, 'CG522-E',  $anchor{'ucgw1'},   	$fmtPurple, 'U' ],#$ci ==  13
		[ $anchor{'tad1_name'}, 'EMG8500', $anchor{'tad1_ip'}, 		     $fmtPurple, 'CF' ]#$ci ==  14
	);

	# Write the output
	my @row = ( '', '', '', '', '', 'Production', '', '', '', $siteid, $workgroup, '', '', '' );
	for ( my $ci = 0; $ci <= $#civals; $ci++ ) { # $ci will be used as parent array's index
		# Cell values
		$row[2]  = $civals[$ci][0];
		$row[1]  = "$row[2].uhc.com";
		$row[4]  = deviceType( $row[2] );
		$row[7]  = $civals[$ci][1];
		$row[8]  = 'Cisco';
		$row[8]  = 'Lantronix' if ( $row[2] =~ /^tad/ ) ;
		$row[12] = $civals[$ci][2];
		if ( $civals[$ci][4] =~ /$sitetype/ and $row[2] =~ /(tad|cis|mls|stk|vgc|wlc)/ ) {
			$rows    = writeRow( $ws, $rows, \@row, $civals[$ci][3] );
		}
		if ( $ci == 7 and $sitetype !~ /^(U)$/ ) {
			# Stacks
			for ( my $ct = 0 ; $ct <= $#StackList ; $ct++ ) {
				my $stack       = $ct + 1;
				my $stackname   = ( split( /,/, $StackList[$ct] ) )[0];
				my $stacksubnet = 'idf' . $stack . '_ds_1';
				$stacksubnet = $anchor{$stacksubnet};
				$row[2]      = $stackname;
				$row[1]      = $row[2] . '.uhc.com';
				$row[4]      = deviceType( $row[2] );
				$row[7]      = $anchor{'stk_model'};
				$row[8]  	 = $anchor{'stack_vendor'};
				$row[12]     = $stacksubnet . '.4';
				$row[12] = $anchor{'loop_subnet'} . '.132' if ( $sitetype =~ /^(N|F)$/ );
				$row[12] = $anchor{'data_subnet_1'} . '.4' if ( $sitetype =~ /^(D|C)$/ );
				$rows        = writeRow( $ws, $rows, \@row, $fmtGreen );
			}
		}
		#for SBC STK
		if ( ($ci == 7) and ($anchor{'sbc'} =~ 'yes') and $sitetype =~ /^(T|G|K|H)$/ ){
			$row[2]      = "stk$anchor{'site_code'}$anchor{'mdf_flr'}a01-sbc";
			$row[1]      = $row[2] . '.uhc.com';
			$row[4]      = 'SBC Switch';
			$row[7]      = 'C9300-24T';
			$row[8]  	 = $anchor{'stack_vendor'};
			$row[12]     = $anchor{'sbc_398_lastip'};
			$rows        = writeRow( $ws, $rows, \@row, $fmtGreen );
		}

		if ($ci == 8 and $anchor{'wlc_nmbr'} == 2 and ($sitetype =~ /^(K|H)$/)) {
			$row[2]  = $anchor{'wlc2_name'};
			$row[1]  = "$row[2].uhc.com";
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'WLC9800';
			$row[8]  = 'Cisco';
			$row[12]     = "$anchor{'loop_subnet'}.39";
			$rows    = writeRow( $ws, $rows, \@row, $civals[$ci][3] );
		}
		if ($ci == 9 and $anchor{'wlc_nmbr'} == 2 and ($sitetype =~ /^(G)$/)){
			$row[2]  = $anchor{'wlc2_name'};
			$row[1]  = "$row[2].uhc.com";
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'WLC9800';
			$row[8]  = 'Cisco';
			$row[12]     = "$anchor{'loop_subnet'}.71";
			$rows    = writeRow( $ws, $rows, \@row, $civals[$ci][3] );
		}
		if ($ci == 10 and $anchor{'wlc_nmbr'} == 2 and ($sitetype =~ /^(T)$/)){
			$row[2]  = $anchor{'wlc2_name'};
			$row[1]  = "$row[2].uhc.com";
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'WLC9800';
			$row[8]  = 'Cisco';
			$row[12]     = "$anchor{'loop_subnet'}.135";
			$rows    = writeRow( $ws, $rows, \@row, $civals[$ci][3] );
		}

		if ( ($ci== 13) and ($sitetype =~ /^(U)$/) and ($anchor{'isp'} eq 'CELLULAR' or $anchor{'pri_provider'} eq 'CELLULAR' ) ){
			$row[2]  = $anchor{'cgw1_name'};
			$row[1]  = "$row[2].uhc.com";
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'CG522-E';
			$row[8]  = 'Cisco';
			$row[12]     = $anchor{'ucgw1'};
			$rows        = writeRow( $ws, $rows, \@row, $fmtBlue );

			if ( $anchor{'isp'} eq 'CELLULAR' and $anchor{'pri_provider'} eq 'CELLULAR' ) {
			$row[2]  = $anchor{'cgw2_name'};
			$row[1]  = "$row[2].uhc.com";
			$row[4]  = deviceType( $row[2] );
			$row[7]  = 'CG522-E';
			$row[8]  = 'Cisco';
			$row[12]     = $anchor{'ucgw2'};
			$rows        = writeRow( $ws, $rows, \@row, $fmtBlue );
			}
		}
	}

	$workbook->close();
	return $outputFile;
}

sub writeEquipmentValidationNew {

	return unless ($anchor{'proj_type'} eq 'build');

	( my $siteid ) = @_;
	my $sitetype = $anchor{'site_type'};
	my $outputFile = $siteid . ' - Equipment Validation Checklist.xls';

	my $workbook = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/$outputFile")
	  or die "create XLS file '$SVR_ROOTDIR/$OutputDir/$outputFile' failed: $!";
	# my $workbook = Spreadsheet::WriteExcel->new("$SVR_ROOTDIR/$OutputDir/Misc/$outputFile")
	#   or die "create XLS file '$SVR_ROOTDIR/$OutputDir/Misc/$outputFile' failed: $!";

	prtout("Writing Equipment Validation Checklist");
	my $ws = $workbook->add_worksheet($siteid);
	$ws->set_zoom(75);
	$ws->set_column( 'A:A', 50 );
		#columns hide all for default
	$ws->set_column('B:B', 15, undef,   1); #tad
	$ws->set_column('C:C', 15, undef,   1); #cis1
	$ws->set_column('D:D', 15, undef,   1); #cis2
	$ws->set_column('E:E', 15, undef,   1); #mls1
	$ws->set_column('F:F', 15, undef,   1); #mls2
	$ws->set_column('G:G', 18, undef,   1); #stack
	$ws->set_column('H:H', 15, undef,   1); #wlc/ap
	$ws->set_column('I:I', 15, undef,   1); #vgc1
	$ws->set_column('J:J', 15, undef,   1); #vgc2


	my $fmtBlank = $workbook->add_format( size => 10 );
	my $fmtRedBlank = $workbook->add_format( size => 10 , color => 'red');
	my $fmtGray = $workbook->add_format(
										 size      => 10,
										 bold      => 1,
										 bg_color  => 22,
										 text_wrap => 1
	);
	my $fmtBlue   = $workbook->add_format( bg_color => 41 );
	my $fmtPurple = $workbook->add_format( bg_color => 31 );
	my $fmtGreen  = $workbook->add_format( bg_color => 42 );
	my $fmtMergeYellow = $workbook->add_format( size => 13, bold => 1, bg_color => 43, align => 'center');
	my $fmtYellow = $workbook->add_format( size => 13, bold => 1, bg_color => 43 );
	my $fmtOrange = $workbook->add_format( bg_color => 47 );
	my $fmtTeal   = $workbook->add_format( bg_color => 35 );
	my $fmtNormal = $workbook->add_format( color => 'black', bold => 0 );
	my $fmtBold = $workbook->add_format( bold => 1 );


	# First three rows are the header
	my $rows   = 1;
	my $height = 12;
	$ws->merge_range( 'A1:J1', "Validate $siteid Equipment Checklist", $fmtMergeYellow );

	my $stack = scalar( @StackList );
	my $stackname = ( split( /,/, $StackList[0] ) )[0];

	my @header = ( '', $anchor{'tad1_name'}, $anchor{'cis1_name'}, $anchor{'cis2_name'},
					$anchor{'mls1_name'}, $anchor{'mls2_name'},	$stackname . ' - ' . $stack,
					$anchor{'wlc1_name'} . ' / wap' . $site_code, $anchor{'vgc1_name'},
					$anchor{'vgc2_name'} );

	$rows = writeRow( $ws, $rows, \@header, $fmtGray );

	my @row = ( 'Login', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'Dial in to test the modem/ssh into Lantronix via LTE IP', '',
			'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A'	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'Ping all other equipment/vlans', 'N/A', '', '', '', '','', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'OS - verify its the standard version', '', '', '', '',
			'', '', 'N/A', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show runn - verify the config', '', '', '', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show cdp neighbor - verify all connections', 'N/A', '',
			'', '', '', '', 'N/A', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show stand br - verify all the interfaces/vlans', 'N/A',
			'N/A', 'N/A', '', '', 'N/A', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show log - look for errors', 'N/A', '', '',
			'', '', '', 'N/A', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show ip int br - verify all the interfaces', 'N/A', '',
			'', '', '', '', 'N/A', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show ip route - verify all routes are present', 'N/A', '',
			'', '', '', 'N/A', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show sw detail - verify # of switches, IOS and priority',
			'N/A', 'N/A', 'N/A', '', '', '', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show power inline - verify all ports have power', 'N/A',
			'N/A', 'N/A', 'N/A', 'N/A', '', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show stack-power - verify the priorities', 'N/A',
			'N/A', 'N/A', '', '', 'N/A', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '**show inventory - verify the network modules, SFPs, Power',
			'N/A', '', '', '', '', '', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '**show env all - verify status of PS, fans, etc.', '', '',
			'', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show int trunk - verify the trunking is correct', 'N/A',
			'N/A', 'N/A', '', '', '', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '**show diag - verify everything has "passed" diagnostics',
			'N/A', '', '', '', '', '', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'verify moh file is there (dir flash:) *only non-SIP sites',
			'N/A', '', '', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show vlan - verify the vlans', 'N/A', 'N/A', 'N/A',
			 '', '', '', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show vtp status - should be transparent', 'N/A', 'N/A',
			'N/A', '', '', '', 'N/A', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show sysinfo - verify version, system information', 'N/A',
			'N/A', 'N/A', 'N/A', 'N/A', 'N/A', '', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show interface summary', 'N/A', 'N/A', 'N/A', 'N/A',
			'N/A', 'N/A', '', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show wlan summary', 'N/A', 'N/A', 'N/A', 'N/A',
			'N/A', 'N/A', '', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show ap summary***', 'N/A', 'N/A', 'N/A', 'N/A',
			'N/A', 'N/A', '', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = ( 'show ap config general <wap name>***', 'N/A', 'N/A',
			'N/A', 'N/A', 'N/A', 'N/A', '', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', 'WLC Specifc ', '','', '','','','', '','' );
	$rows = writeRow( $ws, $rows, \@row, $fmtYellow );

	@row = ( 'show certificate webauth ***', 'N/A', 'N/A', 'N/A',
			'N/A', 'N/A', 'N/A', '', 'N/A', 'N/A'	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = ( 'show network summary (check if webmode is enabled - disable it)',
			'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', '', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'show wlan summary', 'N/A', 'N/A', 'N/A', 'N/A',
			'N/A', 'N/A', '', 'N/A', 'N/A' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '*cis1 notes:', '', '', '', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'Remove bgp network statement so it isnt advertised until cut night',
			 '', '', '', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'Remove bgp network statement so it isnt advertised until cut night',
			'', '', '', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '*stack notes:', '', '', '', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( 'Remove EAC (if applicable)', '', '', '', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '', '', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '**command varies per device type', '',
			'',	'', '', '', '', '', '', '' );
	$rows = writeRow( $ws, $rows, \@row, $fmtBlank );

	@row = ( '** Verify that 2 Power Supplies are present if ordered (cis..01 and 02 routers).',
			'', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = ( '***The APs are usually not staged so that they connect to the WLC.',
			'', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = ( '***So you cannot verify via the WLC, instead ask the staging vendor to take screen shots or somehow give you proof that they staged the APs properly.',
			'', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	@row = ( '***Check if webauth cert is present: if not create one (then reboot) (needed for https)',
			'', '', '', '', '', '', '', '', ''	);
	$rows = writeRow( $ws, $rows, \@row, $fmtRedBlank );

	$ws->set_column('B:B', 15, undef, 0);									#tad
	$ws->set_column('C:C', 15, undef, 0); 									#cis1
	$ws->set_column('D:D', 15, undef, 0) if ( $sitetype !~ /^(C|F)$/ );     #cis2
	$ws->set_column('E:E', 15, undef, 0) if ( $sitetype =~ /^(T|G|K|H)$/ ); #mls1
	$ws->set_column('F:F', 15, undef, 0) if ( $sitetype =~ /^(T|G)$/ );		#mls2
	$ws->set_column('G:G', 15, undef, 0); 									#stack
	$ws->set_column('H:H', 15, undef, 0); 									#wlc/ap
	$ws->set_column('I:I', 15, undef, 0) if ($anchor{'vgcount'} > 0 ); 		#vgc1
	if ($anchor{'vgcount'} == 2 and $sitetype =~ /^(T|G)$/ ) {
		$ws->set_column('J:J', 15, undef, 0); 								#vgc2
	}

	$workbook->close();

	return $outputFile;

}

sub writeSDWANcsvNew {
	prtout( "Writing SDWAN CSV.....");
	return unless ( $anchor{'SDWAN'} eq 'Y' );
	( my ($siteid, $r1 , $r2, $rt_type,$int_type,$int_typer1,$int_typer2,$tport) ) = @_;

	my $int_type_r1 = uc($int_typer1);
	my $int_type_r2 = uc($int_typer2);
	my $outputFile;
	my @PST = ('WA','OR','NV','CA');
	my @MST = ('MT','ID','WY','UT','CO','AZ','NM');
	my @CST = ('ND','SD','MN','WI','NE','IA','IL','KS','MO','KY','OK','AR','TN','TX','LA','MS','AL','MC'); # Mobile clinic (MC) should be CST
	my @EST = ('MI','IN','GA','OH','WV','PA','VA','NC','SC','FL','VT','NH','ME','DC','MD','DE','NJ','CT','MA','RI','NY');
	#my $region = "CST";
	my $site = uc(substr($r1, 3,2));
	my $region;

	if($site ~~ @PST){
		$region = "PST";
	}
	elsif($site ~~ @MST){
		$region = "MST";
	}
	elsif($site ~~ @CST){
		$region = "CST";
	}
	elsif($site ~~ @EST){
		$region = "EST";
	}else{
		$region = "<INTL>";
	}
	prtout( "Region: $region");
	prtout( "router type: $rt_type");
	prtout( "internet type: $int_type");

	my $rtr_type = '';
	my $rtr_type_ng;
	if( $anchor{'tloc'} =~ m/yes/i ){
		$rtr_type = $rt_type;
		if($rt_type =~ m/C8200-1N-4T/i){
			$rtr_type_ng = $rt_type;
			$rtr_type_ng =~ s/^.//;
		}if($rt_type =~ m/C8300-1N1S/){
			$rtr_type_ng = $rt_type;
			$rtr_type_ng =~ s/^.//;
			$rtr_type_ng = $rtr_type_ng . "-6T";
		}if($rt_type =~ m/C8300-2N2S/i){
			$rtr_type_ng = $rt_type;
			$rtr_type_ng =~ s/^.//;
			$rtr_type_ng = $rtr_type_ng . "-4T2X";
		}if($rt_type =~ m/C8500-12X4QC/i){
			$rtr_type_ng = $rt_type;
			$rtr_type_ng =~ s/^.//;
		}
	}else{
		if($rt_type == "4451"){
		$rt_type = '4451-X';
		}
		if (($rt_type eq '4331') or ($rt_type eq '4451-X') or ($rt_type eq '4461')){
		$rtr_type = "ISR" . $rt_type;
		}else{
		$rtr_type = $rt_type;
		}
	}
	#prtout( "RTR type SDWAN CSV: $rtr_type_ng");

	#for secondary router csv template
	my $sec_ckt_type;
	if( $anchor{'pub_provider'} =~ m/Internet/i ){
	$sec_ckt_type = 'INET';
	}else{
	$sec_ckt_type = 'MPLS';
	}

	my $MPLS_provider = '';
	my $MPLS_provider2 = '';
	if (($anchor{'r1_provider'}) eq 'ATT'){
		$MPLS_provider = 'ATT';
	}
	elsif (($anchor{'r1_provider'} eq 'VZ') and ($anchor{'pri_vlan'} > 0)){
		$MPLS_provider = 'VZ';
	}
	elsif (($anchor{'r1_provider'} eq 'VZ') and ($anchor{'pri_vlan'} == 0)){
		$MPLS_provider = "VZ_un";
	}
	elsif (($anchor{'r1_provider'} eq 'LUMEN') and ($anchor{'pri_vlan'} > 0)){
		$MPLS_provider = 'Lumen';
	}
	elsif (($anchor{'r1_provider'} eq 'LUMEN') and ($anchor{'pri_vlan'} == 0)){
		$MPLS_provider = "Lumen_un";
	}
	if (($anchor{'r2_provider'}) eq 'ATT'){
		$MPLS_provider2 = 'ATT';
	}
	elsif (($anchor{'r2_provider'} eq 'VZ') and ($anchor{'r2vlan'} > 0)){
		$MPLS_provider2 = 'VZ';
	}
	elsif (($anchor{'r2_provider'} eq 'VZ') and ($anchor{'r2vlan'} == 0)){
		$MPLS_provider2 = "VZ_un";
	}
	elsif (($anchor{'r2_provider'} eq 'LUMEN') and ($anchor{'r2vlan'} > 0)){
		$MPLS_provider2 = 'Lumen';
	}
	elsif (($anchor{'r2_provider'} eq 'LUMEN') and ($anchor{'r2vlan'} == 0)){
		$MPLS_provider2 = "Lumen_un";
	}
	$anchor{'street'} =~ s/,|\.//g;
	$anchor{'street'} =~ s/ /_/g; #replace spcaes with '_';

	#OLD/Current Gen Folder path
	my $u_folder = "New-Standards/SDWAN/U_templates/Consolidated/";
	my $cf_folder = "New-Standards/SDWAN/CF_templates/Consolidated/";
	my $dn_folder = "New-Standards/SDWAN/DN_templates/Consolidated/";
	my $kh_folder = "New-Standards/SDWAN/KH_templates/Consolidated/";
	my $tg_folder = "New-Standards/SDWAN/TG_templates/Consolidated/";
	my $stg_folder = "New-Standards/SDWAN/Staging/";

	#Next_Gen Folder path
	my $cf_nxtgn_folder = "New-Standards/SDWAN/CF_templates/Next_Gen/";
	my $dn_nxtgn_folder = "New-Standards/SDWAN/DN_templates/Next_Gen/";
	my $kh_nxtgn_folder = "New-Standards/SDWAN/KH_templates/Next_Gen/";
	my $tg_nxtgn_folder = "New-Standards/SDWAN/TG_templates/Next_Gen/";
	
	# SDWAN templates will keep legacy site types in their names for now
	if ( $anchor{'site_type'} =~ /^(C|F)$/ ) {
		if( $anchor{'tloc'} eq 'yes' ){
				if($anchor{'transport'} == 1){
					if($anchor{'int_provider'} =~ m/Internet/i){
						$outputFile = writeTemplate( $cf_nxtgn_folder . "Trust_RS_Consolidated_P_" . $rtr_type_ng . "_INET_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_P_' . $rtr_type_ng . '_INET_dtmpl.csv');
						if ($anchor{'int_type_r1'} eq 'static'){
							$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
						}if ($anchor{'int_type_r1'} eq 'dhcp'){
						$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
						}
					}if($anchor{'r1_provider'} =~ /^(ATT|VZ|LUMEN)$/){
						$outputFile = writeTemplate( $cf_nxtgn_folder . "Trust_RS_Consolidated_P_" . $rtr_type_ng . "_MPLS_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_P_' . $rtr_type_ng . '_MPLS_dtmpl.csv');
						$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Primary_Router_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					}
				}if($anchor{'transport'} == 2){
					if(($anchor{'int_provider'} =~ m/Internet/i) and ($anchor{'pub_provider'} =~ m/Internet/i)){
						$outputFile = writeTemplate( $cf_nxtgn_folder . "Trust_RS_Consolidated_P_" . $rtr_type_ng . "_INET1_INET2_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_P_' . $rtr_type_ng . '_INET1_INET2_dtmpl.csv');
						if ($anchor{'int_type_r1'} eq 'static'){
							$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
						}if ($anchor{'int_type_r1'} eq 'dhcp'){
							$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
						}
					}if(($anchor{'int_provider'} =~ m/Internet/i) and ($anchor{'r1_provider'} =~ /^(ATT|VZ|LUMEN)$/)){
						$outputFile = writeTemplate( $cf_nxtgn_folder . "Trust_RS_Consolidated_P_" . $rtr_type_ng . "_INET_MPLS_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_P_' . $rtr_type_ng . '_INET_MPLS_dtmpl.csv');
						if ($anchor{'int_type_r1'} eq 'static'){
							$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
						}if ($anchor{'int_type_r1'} eq 'dhcp'){
							$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
						}
					}
				}
		}else{
			$outputFile = writeTemplate( $cf_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Pico_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Pico_' . $int_type . '_dtmpl.csv');
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
			}
		}
	}
	if ( $anchor{'site_type'} =~ /^(D|N)$/ ) {
		if( $anchor{'tloc'} =~ m/yes/i ){
			if ($tport == 2){
			$outputFile = writeTemplate( $dn_nxtgn_folder . "Trust_RS_Consolidated_S_" . $rtr_type_ng . "_Pri_INET_1EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_S_' . $rtr_type_ng . '_Pri_INET_1EXT_dtmpl.csv');
			$outputFile = writeTemplate( $dn_nxtgn_folder . "Trust_RS_Consolidated_S_" . $rtr_type_ng . "_Sec_" . $sec_ckt_type . "_1EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_S_' . $rtr_type_ng . '_Sec_' .$sec_ckt_type. '_1EXT_dtmpl.csv');
			if( $anchor{'pub_provider'} =~ m/Internet/i ){
				if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}if( $anchor{'r2_provider'} =~ /^(ATT|VZ|LUMEN)$/ ){
				if($anchor{'int_type_r1'} eq 'static'){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if($anchor{'int_type_r1'} eq 'dhcp'){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}
		}
		if ($tport == 3){
			$outputFile = writeTemplate( $dn_nxtgn_folder . "Trust_RS_Consolidated_S_" . $rtr_type_ng . "_Pri_INET_MPLS_1EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_S_' . $rtr_type_ng . '_Pri_INET_MPLS_1EXT_dtmpl.csv');
			$outputFile = writeTemplate( $dn_nxtgn_folder . "Trust_RS_Consolidated_S_" . $rtr_type_ng . "_Sec_" . $sec_ckt_type . "_2EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_S_' . $rtr_type_ng . '_Sec_' .$sec_ckt_type. '_2EXT_dtmpl.csv');
			if( $anchor{'pub_provider'} =~ m/Internet/i ){
				if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}if( $anchor{'r2_provider'} =~ /^(ATT|VZ|LUMEN)$/ ){
				if($anchor{'int_type_r1'} eq 'static'){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if($anchor{'int_type_r1'} eq 'dhcp'){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}
			# No Available template for 3-transport model... to follow...
		}
		}else{
		if ($tport == 2){
			$outputFile = writeTemplate( $dn_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Small_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Small_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			$outputFile = writeTemplate( $dn_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Small_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region .'_' . $rtr_type . '_Small_Sec_' . $tport . '-Transport_dtmpl.csv' );
		if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}if ($tport == 3){
			$outputFile = writeTemplate( $dn_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Small_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Small_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			$outputFile = writeTemplate( $dn_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Small_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region .'_' . $rtr_type . '_Small_Sec_' . $tport . '-Transport_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}
	}
	}
	if ( $anchor{'site_type'} =~ /^(K|H)$/ ) {
		if( $anchor{'tloc'} eq 'yes' ){
		if ($tport == 2){
			$outputFile = writeTemplate( $kh_nxtgn_folder . "Trust_RS_Consolidated_M_" . $rtr_type_ng . "_Pri_INET_1EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_M_' . $rtr_type_ng . '_Pri_INET_1EXT_dtmpl.csv');
			$outputFile = writeTemplate( $kh_nxtgn_folder . "Trust_RS_Consolidated_M_" . $rtr_type_ng . "_Sec_" . $sec_ckt_type . "_1EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_M_' . $rtr_type_ng . '_Sec_' .$sec_ckt_type. '_1EXT_dtmpl.csv');
			if( $anchor{'pub_provider'} =~ m/Internet/i ){
				if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}if( $anchor{'r2_provider'} =~ /^(ATT|VZ|LUMEN)$/ ){
				if($anchor{'int_type_r1'} eq 'static'){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if($anchor{'int_type_r1'} eq 'dhcp'){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}
		}if ($tport == 3){
			$outputFile = writeTemplate( $kh_nxtgn_folder . "Trust_RS_Consolidated_M_" . $rtr_type_ng . "_Pri_INET_MPLS_1EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_M_' . $rtr_type_ng . '_Pri_INET_MPLS_1EXT_dtmpl.csv');
			$outputFile = writeTemplate( $kh_nxtgn_folder . "Trust_RS_Consolidated_M_" . $rtr_type_ng . "_Sec_" . $sec_ckt_type . "_2EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_M_' . $rtr_type_ng . '_Sec_' .$sec_ckt_type. '_2EXT_dtmpl.csv');
			if( $anchor{'pub_provider'} =~ m/Internet/i ){
				if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}if( $anchor{'r2_provider'} =~ /^(ATT|VZ|LUMEN)$/ ){
				if($anchor{'int_type_r1'} eq 'static'){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if($anchor{'int_type_r1'} eq 'dhcp'){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}
			# No Available template for 3-transport model... to follow...
		}
		}else{
		if ($tport == 2){
			$outputFile = writeTemplate( $kh_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			$outputFile = writeTemplate( $kh_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Sec_' . $tport . '-Transport_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}else{
			$outputFile = writeTemplate( $kh_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Pri_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Pri_' . $int_type . '_dtmpl.csv');
			$outputFile = writeTemplate( $kh_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Medium_Sec_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Medium_Sec_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}
	}
	}
	if ( $anchor{'site_type'} =~ /^(T|G)$/ ) {
		if( $anchor{'tloc'} eq 'yes' ){
		if ($tport == 2){
			$outputFile = writeTemplate( $tg_nxtgn_folder . "Trust_RS_Consolidated_L_" . $rtr_type_ng . "_Pri_INET_1EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_L_' . $rtr_type_ng . '_Pri_INET_1EXT_dtmpl.csv');
			$outputFile = writeTemplate( $tg_nxtgn_folder . "Trust_RS_Consolidated_L_" . $rtr_type_ng . "_Sec_" . $sec_ckt_type . "_1EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_L_' . $rtr_type_ng . '_Sec_' .$sec_ckt_type. '_1EXT_dtmpl.csv');
			if( $anchor{'pub_provider'} =~ m/Internet/i ){
				if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}if( $anchor{'r2_provider'} =~ /^(ATT|VZ|LUMEN)$/ ){
				if($anchor{'int_type_r1'} eq 'static'){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Pri_Static_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if($anchor{'int_type_r1'} eq 'dhcp'){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}
		}if ($tport == 3){
			$outputFile = writeTemplate( $tg_nxtgn_folder . "Trust_RS_Consolidated_L_" . $rtr_type_ng . "_Pri_INET_MPLS_1EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_L_' . $rtr_type_ng . '_Pri_INET_MPLS_1EXT_dtmpl.csv');
			$outputFile = writeTemplate( $tg_nxtgn_folder . "Trust_RS_Consolidated_L_" . $rtr_type_ng . "_Sec_" . $sec_ckt_type . "_2EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_L_' . $rtr_type_ng . '_Sec_' .$sec_ckt_type. '_2EXT_dtmpl.csv');
			if( $anchor{'pub_provider'} =~ m/Internet/i ){
				if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'static')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Sec_Static_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'dhcp')){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . "Sec_DHCP_Internet.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}if( $anchor{'r2_provider'} =~ /^(ATT|VZ|LUMEN)$/ ){
				if($anchor{'int_type_r1'} eq 'static'){
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}if($anchor{'int_type_r1'} eq 'dhcp'){
					$outputFile = writeTemplate( $stg_folder . "Pri_DHCP_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
					$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
				}
			}
			# No Available template for 3-transport model... to follow...
		}if ($tport == 4){
			$outputFile = writeTemplate( $tg_nxtgn_folder . "Trust_RS_Consolidated_L_" . $rtr_type_ng . "_Pri_INET_MPLS_2EXT_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_L_' . $rtr_type_ng . '_Pri_INET_MPLS_2EXT_dtmpl.csv');
			$outputFile = writeTemplate( $tg_nxtgn_folder . "Trust_RS_Consolidated_L_" . $rtr_type_ng . "_Sec_INET_MPLS_2EXT_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_L_' . $rtr_type_ng . '_Sec_INET_MPLS_2EXT_dtmpl.csv');
			if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'static')){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if (($anchor{'int_type_r1'} eq 'static') and ($anchor{'int_type_r2'} eq 'dhcp')){
				# No Available template for 4-transport model... to follow...
			    prtout( "\n NO AVAILABLE SDWAN TEMPLATE and BOOTSTRAP FOR NEW $tport-TRANSPORT DESIGN. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
			}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'static')){
				# No Available template for 4-transport model... to follow...
			    prtout( "\n NO AVAILABLE SDWAN TEMPLATE and BOOTSTRAP FOR NEW $tport-TRANSPORT DESIGN. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
			}if (($anchor{'int_type_r1'} eq 'dhcp') and ($anchor{'int_type_r2'} eq 'dhcp')){
			# No Available template for 4-transport model... to follow...
			prtout( "\n NO AVAILABLE SDWAN TEMPLATE and BOOTSTRAP FOR NEW $tport-TRANSPORT DESIGN. PLEASE CONTACT YOUR CONFIGURATOR ADMIN!\n");
			}
		}
		}else{
		if ($tport == 2){
			$outputFile = writeTemplate( $tg_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Pri_" . $int_type . "_" . $tport . "-Transport_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Pri_' . $int_type . '_' . $tport . '-Transport_dtmpl.csv');
			$outputFile = writeTemplate( $tg_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Sec_" . $tport . "-Transport_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Sec_' . $tport . '-Transport_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_NO_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}else{
			$outputFile = writeTemplate( $tg_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Pri_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Pri_' . $int_type . '_dtmpl.csv');
			$outputFile = writeTemplate( $tg_folder . "Trust_RS_Consolidated_" . $rtr_type . "_Large_Sec_dtmpl.csv", $r2 . ' - Trust_RS_' . $region . '_' . $rtr_type . '_Large_Sec_dtmpl.csv' );
			if ($int_type eq 'static'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Static_Internet_and_" . $MPLS_provider . "_MPLS.txt", $r1 . ' - Staging_Router_Configuration.txt');
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}if ($int_type eq 'dhcp'){
				$outputFile = writeTemplate( $stg_folder . $rtr_type . "_Secondary_Router_" . $MPLS_provider2 . "_MPLS.txt", $r2 . ' - Staging_Router_Configuration.txt');
			}
		}
	}
	}
	if ( $anchor{'site_type'} =~ /^(U)$/ ) {
		$outputFile = writeTemplate( $u_folder . "Trust_MC_Consolidated_" . $rtr_type . "_" . $int_type . "_dtmpl.csv", $r1 . ' - Trust_MC_' . $region . '_' . $rtr_type . "_" . $int_type . '_dtmpl.csv');
		$outputFile = writeTemplate( $stg_folder . $rtr_type . "_DHCP_Internet.txt", $r1 . ' - Staging_Router_Configuration.txt'); #DHCP template for Micro site staging
	}

	return $outputFile;
}

} #This curly braket is the end of the else for new site models

###################################################################################
#************************SUBS FOR NEW SITE MODELS END HERE************************#
###################################################################################




##################################################################################
#*****************************GLOBAL SUBS START HERE*****************************#
##################################################################################

sub setAnchorValuesGlobal {
	my ( $id, $site_code, $environ );
	( $id, $site_code, $environ ) = @_;
	$anchor{'site_code'} = $site_code;

	# Script version
	$anchor{'confver'}     = $VERSION;
	$anchor{'uhgwireless'} = '';

	# Form data
	# 05-24 Update: Comment out anchors has been moved to the new ASN tables
	my $query = $dbh->prepare("SELECT * FROM nslanwan.configurator_forms where id='$id'")
	  or die "Query of form information failed: " . $dbh->errstr;
	$query->execute;
	while ( my @form = $query->fetchrow_array ) {

		$anchor{'pri_e_pvc'}          = $form[5];
		$anchor{'mci_gold_car'}       = $form[9];
		$anchor{'loop_oct2'}          = $form[10];
		$anchor{'loop_oct3'}          = $form[11];
		$anchor{'site_eng_phone'}     = $form[13];
		$anchor{'drawn_by_phone'}     = $form[13];
		$anchor{'site_contact'}       = $form[14];
		$anchor{'site_contact_phone'} = $form[15];
		$anchor{'site_type'}          = $form[16];
		$anchor{'pri_circuit_type'}   = $form[17];
		$anchor{'region'}             = $form[18];
		$anchor{'fw'}                 = $form[19];
		$anchor{'ips'}                = $form[20];
		$anchor{'ipt'}                = $form[21];
		$anchor{'wlan'}               = $form[22];
		$anchor{'SDWAN'}         	  = $form[23];
		$anchor{'phase'}              = $form[24];
		$anchor{'xsubnet'}            = $form[25];
		$anchor{'mdf_bldg'}           = $form[26];
		$anchor{'mdf_flrnumber'}      = $form[27];
		$anchor{'proj_id'}					= $form[33];
		$anchor{'Infrasubnet'}      = $form[34];
		$anchor{'subrate'}            = $form[35];
		$anchor{'telepresence'}       = $form[37];
		$anchor{'uhgdivision'}        = $form[42];
		$anchor{'router_seltype'}     = $form[43];
		$anchor{'vgcount'}            = $form[44];
		$anchor{'man_campus'}         = $form[45];
		$anchor{'proj_type'}          = $form[46];
		$anchor{'upgrade_sel'}         = $form[48];
		$anchor{'latitude'}          = $form[51];
		$anchor{'longitude'}          = $form[52];
		$anchor{'site_id'}          = $form[53];
		$anchor{'sudi_sn'}          = $form[54];
		$anchor{'int_bw'}          = $form[57];
		$anchor{'sudi_sn2'}          = $form[58];
		$anchor{'int_type'}          = $form[61];
		$anchor{'loop_oct1'}          = $form[63];
		$anchor{'wlc_model'}					  = $form[67];
		$anchor{'wlc_nmbr'}					  = $form[68];
	}

	#stack vendor
	my $query = $dbh->prepare("SELECT * FROM nslanwan.configurator_stacks where config_id='$id' and mailRoute='$site_code'")
		or die "Query of form information failed: " . $dbh->errstr;
	$query->execute;
	while ( my @form2 = $query->fetchrow_array ) {
		$anchor{'stack_vendor'}       = $form2[9];
	}

	#PrimaryBU
	my $query = $dbh->prepare("SELECT p.primarybu FROM nslanwan.projects p where p.mailRoute ='$site_code' and p.proj_id = '$anchor{'proj_id'}'")
		or die "Query of form information failed: " . $dbh->errstr;
	$query->execute;
	while ( my @form3 = $query->fetchrow_array ) {
		$anchor{'primary_bu'}       = $form3[0];
	}

	#BOT API table query
	my @apiindx;
	my $apimsg;
	my $proj_id = $anchor{'proj_id'};
	my $tblservices_header = ['Content-Type' => 'application/json; charset=UTF-8', 'Authorization' => 'Basic ' . encode_base64($SMB_USER . ':' . $SMB_PASS)];
	my @apitbl = (
		"table=nslanwan.tblservices&criteria=(services like ".uri_escape("'%,sbc,%'")." or services like ".uri_escape("'sbc,%'")." or services like".uri_escape("'%,sbc'").") and mailRoute='$site_code'",
		"criteria=(proj_id = ".uri_escape("'$proj_id'").") AND (mailRoute = '$site_code') AND (tloc = 'yes')&table=nslanwan.projects"
	);
	foreach my $apiquery (@apitbl){
		$api_url = "$api_url$apiquery";
		my $request = HTTP::Request->new(GET =>$api_url ,$tblservices_header);
		my $ua = LWP::UserAgent->new;
		my $response = $ua->request($request);
		my $apioutput = $response->decoded_content;
		if ($apioutput  =~ 'errorMessage'){
		$apimsg = 'no';
		}else{
		$apimsg = 'yes';
		}
		push @apiindx, $apimsg;
		my @api_combine = split(/\?/,$api_url);
        $api_url = "@api_combine[0]?";
	}
	$anchor{'sbc'} = @apiindx[0];
	$anchor{'tloc'} = @apiindx[1];
	prtout("api_sbc: $anchor{'sbc'}");
	prtout("api_tloc: $anchor{'tloc'}");

	#ASN Tables
	my $r1_provider			= '';
	my $r2_provider			= '';

	my $query = $dbh->prepare("SELECT x.* FROM nslanwan.site_asn x WHERE asnStatus ='Active' AND status IS NOT NULL AND mailRoute ='$site_code'")
		or die "Query of form information failed: " . $dbh->errstr;
	$query->execute;
	while ( my @form4 = $query->fetchrow_array ) {
		$anchor{'transport'}          = $form4[2];
	if($form4[10] eq 'R1-P'){
        $anchor{'int_cer_ip'}    		= $form4[6];
        $anchor{'int_per_ip'}    		= $form4[7];
		$anchor{'int_ckt_id'}			= $form4[13];
		$anchor{'int_provider'}			= $form4[3];
		$anchor{'int_type_r1'}          = $form4[16];
    }
    if($form4[10] eq 'R1-S'){
        $anchor{'pri_wan_ip_cer'}     	= $form4[6];
        $anchor{'pri_wan_ip_per'}     	= $form4[7];
		$anchor{'pri_ckt_id'}			= $form4[13];
		$anchor{'pri_vlan'}				= $form4[9];
		$anchor{'bgp_as'}             = $form4[4];
        $r1_provider       				= $form4[3];
		if($r1_provider =~ m/AT\&T/i ){
			$anchor{'r1_provider'} = 'att';
		}
		elsif($r1_provider =~ m/Verizon/i ){
			$anchor{'r1_provider'} = 'vz';
		}
		elsif($r1_provider =~ m/LUMEN/i ){
			$anchor{'r1_provider'} = 'lumen';
		}
		elsif($r1_provider =~ m/MPLS Cloud/i ){
			$anchor{'r1_provider'} = 'att';
		}
	}
    if(($form4[10] eq 'R2-P') or ($form4[10] eq 'R1-T')){
        $anchor{'pub_cer_ip'}     		= $form4[6];
        $anchor{'pub_per_ip'}    		= $form4[7];
		$anchor{'pub_ckt_id'}			= $form4[13];
		$anchor{'pub_provider'}			= $form4[3];
		$anchor{'int_type_r2'}          = $form4[16];
    }
    if($form4[10] eq 'R2-S'){
        $anchor{'sec_wan_ip_cer'}     	= $form4[6];
        $anchor{'sec_wan_ip_per'} 		= $form4[7];
		$anchor{'sec_ckt_id'}			= $form4[13];
		$anchor{'r2vlan'}				= $form4[9];
		$anchor{'bgp_as'}             = $form4[4];
        $r2_provider       				= $form4[3];
		if($r2_provider =~ m/AT\&T/i ){
			$anchor{'r2_provider'} = 'att';
		}
		elsif($r2_provider =~ m/Verizon/i ){
			$anchor{'r2_provider'} = 'vz';
		}
		elsif($r2_provider =~ m/LUMEN/i ){
			$anchor{'r2_provider'} = 'lumen';
		}
		elsif($r2_provider =~ m/MPLS Cloud/i ){
			$anchor{'r2_provider'} = 'att';
		}
    }

	}
	prtout( "BGP ASN: $anchor{'bgp_as'}\n");

	#internet bandwidth subrate
	my ($dnbw,$upbw) = split(/\//, $anchor{'int_bw'});
	if ($upbw == $dnbw){
		$anchor{'isp_subrate'} = $dnbw * .85;
	}else{
		$anchor{'isp_subrate'} = $upbw;
	}
	$anchor{'upbw'} = $upbw / 1000;
	$anchor{'dnbw'} = $dnbw / 1000;

	# Hardcoded wireless radius settings per SRT15935059 (email from Deven/Steve said to hard code these)
	$anchor{'wireless_radius1'}     = 10;
	$anchor{'wireless_radius2'}     = 11;
	$anchor{'wireless_radius3'}     = 12;
	$anchor{'wireless_radius4'}     = 13;
	prtout( "Stack Vendor: " .$anchor{'stack_vendor'});
	prtout( "Project Type: " . $anchor{'proj_type'} );
	prtout( "UHG Division Selected: " . $anchor{'uhgdivision'} );
	prtout( "User Selected Router Type: " . $anchor{'router_seltype'} );
	prtout( "Internet Type: " . $anchor{'int_type'} );
	prtout( "R1 Internet Type: " . $anchor{'int_type_r1'} );
	prtout( "R2 Internet Type: " . $anchor{'int_type_r2'} );
	prtout( "Transport: " . $anchor{'transport'} );
	prtout( "MAN Campus: " . $anchor{'man_campus'} );



	# Location information
	$query = $dbh->prepare(
			 "SELECT Address, City, State, Zipcode, SiteEngineer, sitemodel, modemNumber FROM $DB_NSLCM.siteList where mailRoute like '%$site_code%'" )
	  or die "Query of location information failed: " . $dbh->errstr;
	$query->execute;
	while ( my @location = $query->fetchrow_array ) {
		$anchor{'street'}     = $location[0];
		$anchor{'city'}       = $location[1];
		$anchor{'state'}      = $location[2];
		$anchor{'zip'}        = $location[3];
		$anchor{'site_eng'}   = $location[4];
		$anchor{'drawn_by'}   = $location[4];
		$anchor{'site_model'} = $location[5];
		$anchor{'oob_tad'} = $location[6];
	}
	if ( $anchor{'region'} eq 'USA' ) {
	  my @wireless_east =('AL','CT','DC','DE','FL','GA','IL','IN','KY','MA','MD','ME','MI','MS','NC','NH','NJ','NY','OH','PA','RI','SC','TN','VA','VT','WI','WV');
    my @wireless_west = ('AK','AR','AZ','CA','CO','HI','IA','ID','KS','LA','MN','MO','MT','ND','NE','NM','NV','OK','OR','SD','TX','UT','WA','WY');

    my $state = $anchor{'state'};
    $state =~ tr/a-z/A-Z/;
		if($state ~~ @wireless_east){
		  $anchor{'wireless_region'} = 'EAST';
		}
		elsif($state ~~ @wireless_west){
		  $anchor{'wireless_region'} = 'WEST';
		}
		else {
			$anchor{'wireless_region'} = 'INTL';
		}
	}
	$sitemodel = $anchor{'site_model'};


	# Differentiate between Large sites with /18 and /19 address space
	my $sitetype = $anchor{'site_type'};    # this is only used for a print statement below
	if ( $anchor{'site_type'} eq 'L/18' ) {
		$anchor{'site_type'} = 'X';
	} elsif ( $anchor{'site_type'} eq 'L/19' ) {
		$anchor{'site_type'} = 'L';
	}

	# Do initial input check
	inputErrorCheckGlobal();

	# File pathing is based on the division and environment
	our $EnvPath;
	if ( $anchor{'uhgdivision'} eq 'UHG' ) {
		$EnvPath = 'web';                               # defauwlt to prod
		$EnvPath = 'webdev' if ( $environ eq 'dev' );
	} elsif ( $anchor{'uhgdivision'} eq 'TRICARE' ) {
		$EnvPath = 'tcw';                               # default to prod
		$EnvPath = 'tcwdev' if ( $environ eq 'dev' );
	}
	our $SMBBase = 'SHARED/UHTServerOP/LAN_WAN_MAN/TOOLS/Config';
	$SMB_TEMPLATE_DIR = "$SMBBase/$EnvPath/templates";
	$SMB_FIN_DIR      = "$SMBBase/$EnvPath/finished";
	$SMB_ROOTDIR      = "Unpiox56pn/netsvcs/$SMBBase/$EnvPath";
	$ROOTDIR          = nslanwanfilepath() . '/scripts';          # external routine from require file
	$SVR_ROOTDIR 	  = nslanwanfilepath() . '/htdocs/tmp';

	# Output directory name
	( my $day, my $month ) = ( localtime(time) )[ 3, 4 ];
	$month++;
	$OutputDir = $site_code . "-$month-$day";                     # Variable is a global declared prior to calling this sub
	                                                              # prtout("Output will be in /$SMB_ROOTDIR/finished/$OutputDir");
	prtout( '', "Identified site code $site_code for input parameters using form ID #$id and site type $sitetype" );

	# Modify Phase value
	if ( $anchor{'phase'} eq "Phase 1 - MPLS Untrusted" ) {
		$anchor{'phase'} = 'MPLSUntrusted';
		prtout( "Identified Build Phase: $anchor{'phase'}", '' );
	} elsif ( $anchor{'phase'} eq 'Phase 1 - MPLS Trusted' ) {
		$anchor{'phase'} = 'MPLSTrusted';
		prtout( "Identified Build Phase: $anchor{'phase'}", '' );
	} elsif ( $anchor{'phase'} eq 'Fully Trusted' ) {
		$anchor{'phase'} = 'FullyTrusted';
		prtout( "Identified Build Phase: $anchor{'phase'}", '' );
	} else {
		prtout( "Unidentified build Phase: '$anchor{'phase'}'", '' );
	}

	# QOS policies
	my $wantype = $anchor{'pri_circuit_type'};
	my $subfull = $anchor{'subrate'} * 1000;
	if ( $wantype eq 'DS3_subrate' ) {
		prtout( "WANTYPE is $wantype", "DS3 Subraate Port Speed detected: $subfull" );
		$anchor{'sec_circuit_type'} = $wantype;
		$anchor{'subrate_full'}     = $subfull;
		$anchor{'backup_police'}    = $subfull * .15;
		prtout("$anchor{'backup_police'} is equal to $subfull times 15%");
	} elsif ( $wantype eq 'Metro_Ethernet' or $wantype eq 'MPLS_Ethernet' ) {
		prtout( "WANTYPE is $wantype", "Ethernet Port Speed detected is $subfull" );
		$anchor{'sec_circuit_type'} = $wantype;
		$anchor{'subrate_full'}     = $subfull;
		$anchor{'backup_police'}    = $subfull * .15;
		prtout("$anchor{'backup_police'} is equal to $subfull times 15%");
	} else {
		$anchor{'sec_circuit_type'} = $wantype;
		$query = $dbh->prepare(
"SELECT port_speed, backup_speed FROM cktspeedmappings,backup_policy WHERE ckt_type='$wantype' and cktspeedmappings.id=backup_policy.ckt_type_id"
		) or die "Query of port speed mappings failed: " . $dbh->errstr;
		$query->execute;
		while ( my @QOS = $query->fetchrow_array ) {
			$anchor{'backup_police'} = $QOS[1];
		}
	}

	# QOS Queue limit numbers, X, Y, Z and the shaping value
	my $xfactor = 15;
	my $subrate = $anchor{'subrate'};
	if ( $subrate >= 10000 and $subrate < 15000 ) {
		$xfactor = 10;
	} elsif ( $subrate >= 15000 and $subrate < 20000 ) {
		$xfactor = 8;
	} elsif ( $subrate >= 20000 and $subrate < 25000 ) {
		$xfactor = 7;
	} elsif ( $subrate >= 25000 and $subrate < 1000000 ) {
		$xfactor = 6;
	}
	$anchor{'xqueue'}        = int( $subrate / 1000 * $xfactor );
	$anchor{'yqueue'}        = $anchor{'xqueue'};
	$anchor{'zqueue'}        = $anchor{'xqueue'} * 2;
	$anchor{'subrate_shape'} = int( $subrate * .85 );
	$anchor{'desc_fullrate'} = int( $subrate / 1000 );
	my $desc_subrate = ( $subrate / 1000 * .85 );
	
	#! Manipulate subrate values with float variable
	if($desc_subrate =~ /-*\d+\.\d+/){
		$anchor{'desc_subrate'} = sprintf("%.2f", "$desc_subrate");
	}else{
		$anchor{'desc_subrate'} = $desc_subrate;
	}

	# Voice Priorities
	$anchor{'voice_priority'}    = floor( $anchor{'mci_gold_car'} );
	$anchor{'voice_cac'}         = floor( $anchor{'voice_priority'} / 30 + .5 ) * 24;
	$anchor{'high_wan_priority'} = floor( $anchor{'mci_gold_car'} / .6 * .24 );
	$anchor{'low_wan_priority'}  = floor( $anchor{'mci_gold_car'} / .6 * .12 );

	# Region data
	my $state = '';
	if ( $anchor{'region'} eq 'USA' ) {
		prtout("Region: USA");
		$state = $anchor{'state'};
		$state =~ tr/a-z/A-Z/;
		if ( !( defined $hostnetflow{$state} ) or $hostnetflow{$state} eq '' ) {
			prtout(
					"There appears to be a problem in looking up host information for the state selected below.",
					"State: $state",
					"Please verify the state abbreviation is correct on the NSLANWAN website and rerun Configurator.",
					"If there continues to be issues, please contact Steve or Dan."
			);
			xit(1);
		}
	} else {
		prtout("Region: non-USA: $anchor{'region'}");
		$state = $anchor{'region'};
	}
	#prtout("State: $state");
	$anchor{'netflow_host'} = $hostnetflow{$state};


# Get the WLC Radius IP addresses for the state
	$query = $dbh->prepare(
"SELECT configurator_wlc.wlc_radius_1, configurator_wlc.wlc_radius_2, configurator_wlc.wlc_radius_3 from configurator_wlc, configurator_location where configurator_location.locale_id='$state' and configurator_location.wlc_group=configurator_wlc.id"
	) or die "Query of WLC Radius IP address mapping failed: " . $dbh->errstr;
	$query->execute();
	while ( my @ret = $query->fetchrow_array ) {
		$anchor{wlc_radius_1} = $ret[0];
		$anchor{wlc_radius_2} = $ret[1];
		$anchor{wlc_radius_3} = $ret[2];
	}

# Set the country code which is needed for 9800 WLC configs
	if (substr($anchor{site_code}, 0, 2) =~ /^(PH|UK|CB|AU|IR)$/){
		$anchor{'country_code'} = substr($anchor{site_code}, 0, 2);
	}if(substr($anchor{site_code}, 0, 2) =~ /^(II)$/){
		$anchor{'country_code'} = 'IN';
	}else{
		$anchor{'country_code'} = 'US';
	}

	# Telepresence - this was split into US/non-US in the old code but uses the same values
	my %telepresence = (
						 'None',    100,   'PolycomSin', 1000,  'PolycomMul', 3000,  'CTS500',  6000,  'CTS1000', 6000,
						 'CTS3000', 16000, 'BW1000',     1000,  'BW2000',     2000,  'BW3000',  3000,  'BW4000',  4000,
						 'BW5000',  5000,  'BW6000',     6000,  'BW7000',     7000,  'BW8000',  8000,  'BW9000',  9000,
						 'BW10000', 10000, 'BW11000',    11000, 'BW12000',    12000, 'BW13000', 13000, 'BW14000', 14000,
						 'BW15000', 15000, 'BW16000',    16000,
	);
	my $tp = $anchor{'telepresence'};
	$anchor{'ivideo'} = $telepresence{$tp};
	prtout("Telepresence: $tp  BW: $anchor{'ivideo'}");

	# SNMP for TAD (temporary..to be updated)
	( $anchor{'snmp_host1'}, $anchor{'snmp_host2'}, $anchor{'snmp_host3'}, $anchor{'snmp_host4'} ) =
	  split( /,/, $hostsnmp{$state} );

	#Dynamic SNMP for STK,MLS,VGC
	my @snmp_con = ();
	my @snmp_con_aruba = ();
	my @snmp_con_vgc = ();
	my @snmp_con_mkh = ();
	my @snmp_con_xltg = ();
	my @snmp_con_wlc = ();
	my @snmp_list = split( /,/, $hostsnmp{$state} );
		foreach my $snmp(@snmp_list){
		$anchor{'snmp_host_ip'} = $snmp;
		if ($anchor{'stack_vendor'} eq 'aruba'){
			my $snmp_temp_aruba = smbReadFile("Modules/aruba_snmp_config.txt");
			push @snmp_con_aruba, $snmp_temp_aruba;
		}else{
			my $snmp_temp_cisco = smbReadFile("Modules/snmp_config.txt");
			push @snmp_con, $snmp_temp_cisco;
		}
		my $snmp_temp_vgc = smbReadFile("Modules/vgc_snmp_config.txt");
		push @snmp_con_vgc, $snmp_temp_vgc;
		my $snmp_temp_mkh = smbReadFile("Modules/mkh_snmp_config.txt");
		push @snmp_con_mkh, $snmp_temp_mkh;
		my $snmp_temp_xltg = smbReadFile("Modules/xltg_snmp_config.txt");
		push @snmp_con_xltg, $snmp_temp_xltg;
		my $snmp_temp_wlc = smbReadFile("Modules/wlc_snmp_config.txt");
		push @snmp_con_wlc, $snmp_temp_wlc;
		}
	$anchor{'snmp_config'} = join "\r\n", @snmp_con;
	$anchor{'aruba_snmp_config'} = join "\r\n", @snmp_con_aruba;
	$anchor{'vgc_snmp_config'} = join "\r\n", @snmp_con_vgc;
	$anchor{'mkh_snmp_config'} = join "\r\n", @snmp_con_mkh;
	$anchor{'xltg_snmp_config'} = join "\r\n", @snmp_con_xltg;
	$anchor{'wlc_snmp_config'} = join "\r\n", @snmp_con_wlc;

	# DNS
	my $dnsi = 0;
	$anchor{'dns_host'} = '';
	foreach ( split( /,/, $hostdns{$state} ) ) {
		if ($anchor{'stack_vendor'} eq 'aruba'){
		$anchor{'dns_host'} .= "ip dns server-address $_\r\n";
		$dnsi = $dnsi + 1;
		}else{
		$anchor{'dns_host'} .= "ip name-server $_\r\n";
		$dnsi = $dnsi + 1;
		}
		$anchor{'dns_host_tad'} .= "set network dns $dnsi ipaddr $_\r\n";
	}
	$anchor{'dns_host'} .= '!';    # need to manually add a bang at the end

	# DHCP
	$anchor{'dhcp_host'}       = '';
	$anchor{'dhcp_host_nexus'} = '';
	my $ct = 1;
	foreach ( split( /,/, $hostdhcp{$state} ) ) {
		if (($ct >= 1) and ($ct <= 4)){
			$anchor{'dhcp_host'}       .= " ip helper-address $_\r\n";        # end with CR+LF for Visio output
			$anchor{'dhcp_host_nexus'} .= " ip dhcp relay address $_\r\n";    # ditto
			my $dhcphash = 'dhcp_host' . $ct;
			$ct++;
			$anchor{$dhcphash} = $_;
		}
		elsif (($anchor{'man_campus'} eq 'Sierra Nevada') and ($ct >= 4)){
			$anchor{'dhcp_host'}       .= " ip helper-address $_\r\n";        # end with CR+LF for Visio output
			$anchor{'dhcp_host_nexus'} .= " ip dhcp relay address $_\r\n";    # ditto
			my $dhcphash = 'dhcp_host' . $ct;
			$ct++;
			$anchor{$dhcphash} = $_;
		}
	}

	# IPT, Radius, Logging, ATT ASN

	#Updated radius list of ISE servers for STKs
	my %radius_names = ( '10.135.56.158',		'PH518_ISE_Default',
						 '10.195.216.66',	'II747_ISE_Default',
						 '10.73.27.10',		'IR777_ISE_Default',
						 '10.130.130.55',	'II757_ISE_Default',
						 '10.50.48.7',		'ELR_ISE_Default',
						 '10.114.144.8',	'PLY_ISE_Default',
						 '10.203.36.7',		'CTC_ISE_Default'
	);

	#
	( $anchor{'ipt_host1'}, $anchor{'ipt_host2'} ) =
	  split( /,/, $hostipt{$state} );

	#( $anchor{'radius_host1'},  $anchor{'radius_host2'},  $anchor{'radius_host3'} )  = split( /,/, $hostradius{$state} );

	my @radius_con = ();
	my @radius_con2 = ();
	my $uhg_rad_name = ();
	my @uhg_grp_rad = ();
	$anchor{'uhg_radius_group'} = '';

	my @radius = split( /,/, $hostradius{$state});
		foreach my $rad(@radius){
			$anchor{'radius_host'} = $rad;
			$anchor{'radius_host_name'} = $radius_names{$anchor{'radius_host'}};

			my $temp_con = smbReadFile("Modules/stk_radius.txt");
			$temp_con = smbReadFile("Modules/aruba_stk_radius.txt")
			if ( $anchor{'stack_vendor'} eq 'aruba' );

			my $temp_con2 = smbReadFile("Modules/aruba_stk_svr_aaa.txt");
			push @radius_con, $temp_con;
			push @radius_con2, $temp_con2;

			#manipulation for uhg_radius_group anchors
			$uhg_rad_name = $anchor{'radius_host_name'};
			#$anchor{'uhg_radius_group'} = "server name $uhg_rad_name\n";
			my $uhg_grp_line = "server name $uhg_rad_name";
			push @uhg_grp_rad, $uhg_grp_line;
		}

	$anchor{'stk_radius'} = join "\n", @radius_con;
	$anchor{'stk_svr_aaa'} = join "\n", @radius_con2;
	$anchor{'uhg_radius_group'} = join "\n", @uhg_grp_rad;

	#Special radius servers for PH sites

	#9800 Radius server RAD_EAC_ISE
	my @radius_con_9800 = ();
	my %eac_vip = ( '10.135.56.158',		'PH518_EAC_VIP',
						 '10.195.216.66',	'II747_EAC_VIP',
						 '10.73.27.10',		'IR777_EAC_VIP',
						 '10.130.130.55',	'II757_EAC_VIP',
						 '10.50.48.7',		'ELR_EAC_VIP',
						 '10.114.144.8',	'PLY_EAC_VIP',
						 '10.203.36.7',		'CTC_EAC_VIP'
	);
	my @radius_9800 = split( /,/, $hostradius{$state});
		foreach my $rad_wlc(@radius_9800){
			$anchor{'radius_host_ip'} = $rad_wlc;
			$anchor{'9800_eac_vip'} = $eac_vip{$anchor{'radius_host_ip'}};
			my $temp_con_9800 = smbReadFile("New-Standards/Wireless/9800_radius.txt");
			push @radius_con_9800, $temp_con_9800;
		}
	$anchor{'rad_eac_vips'} = join "\n", @radius_con_9800;

	my @logging_con = ();
	my $ctlog = 1;
	my @log_list = split( /,/, $hostlogging{$state} );
		foreach my $log(@log_list){
			$anchor{'logging_host_ip'} = $log;
			my $log_temp = smbReadFile("Modules/stk_logging.txt");
			push @logging_con, $log_temp;

			#tad Logging
			$anchor{'tad_logging'} = "set services syslogserver$ctlog $log";
			$ctlog++;
		}
	$anchor{'stk_logging'} = join "\n", @logging_con;



	$anchor{'att_asn'} = $attasn{$state};

	# Interface type
	my $mpls_upl_int;
	my %interfaceType = (
						  'DS3',            [ 'S1/0',        's1-0',       'S1/0' ],
						  'DS3_subrate',    [ 'S1/0',        's1-0',       'S1/0' ],
						  'IMA',            [ 'ATM0/0.150',  'atm0-0-150', 'ATM0/0.150' ],
						  'IMA-E1',         [ 'ATM0/0.150',  'atm0-0-150', 'ATM0/0.150' ],
						  'MLPPP',          [ 'Multilink99', 'mu99',       'Multilink99' ],
						  'MLPPP-E1',       [ 'Multilink99', 'mu99',       'Multilink99' ],
						  'GLBP',           [ 'S0/0/0:0',    's0-0-0',     'S0/0/0:0' ],
						  'MPLS_Ethernet',  [ "Gi0/0/$mpls_upl_int",     "g0-0-$mpls_upl_int",     "Gi0/0/$mpls_upl_int" ],
						  'Metro_Ethernet', [ 'Gi0/1',       'g0-1',       'Gi0/1' ],
						  'default',        [ 'S0/0/0:0',    's0-0-0',     'S0/0/0:0' ],
	);

	# Default values
	$anchor{'att_upl_int'}     = 'S0/0/0:0';
	$anchor{'att_upl_int_dns'} = 's0-0-0';
	$anchor{'att_upl_int'}     = 'S0/0/0:0';

	#for upgrade only define the encapsulation
	if($anchor{'proj_type'} eq 'upgrade'){
		$anchor{'att_vlan'} = ('encapsulation dot1Q ' . $anchor{'pri_vlan'});
	}else{
		$anchor{'att_vlan'} = $anchor{'pri_vlan'};
	}

	my $tmpwantype = $wantype;
	if ( $tmpwantype =~ /(MLPPP-?E?1?|IMA-?E?1?)/ ) {
		$tmpwantype = $1;
	}
	if ( defined $interfaceType{$tmpwantype} ) {
		my @val = @{ $interfaceType{$tmpwantype} };
		$anchor{'mpls_upl_int'}     = $val[0];
		$anchor{'att_upl_int_dns'} = $val[1];
		$anchor{'mci_upl_int'}     = $val[2];
	}
	$anchor{'sec_e_pvc'} = $anchor{'mci_gold_car'};

	prtout("Circuit Type: $anchor{'pri_circuit_type'}");
	if ( $anchor{'pri_circuit_type'} =~ /^(Metro_Ethernet|MPLS_Ethernet|DS3_subrate)$/ ) {
		prtout("Port Speed: $anchor{'subrate'}");
	}
	prtout("Primary Router Type: $anchor{'router_seltype'}");

	#For Interface selection
	if ( $wantype eq 'MPLS_Ethernet' ) {
		if( $anchor{'router_seltype'} eq 'C8200-1N-4T' ){
			$mpls_upl_int = 2;
			$anchor{'mpls_upl_int'}     = "Gi0/0/$mpls_upl_int";
			$anchor{'att_upl_int_dns'} = "g0-0-$mpls_upl_int";
			$anchor{'mci_upl_int'}     = "Gi0/0/$mpls_upl_int";
			$anchor{'isp_upl_int'}     = 'Gi0/0/3';
			$anchor{'pub_upl_int'}     = 'Gi0/0/3';

			#Below will cover 4-transport and Dual-DIA align to the next-Gen Arch design.
			if( $anchor{'tloc'} eq 'yes' ){
				if($anchor{'transport'} == 2){
					$anchor{'tloc_int'}     = 'Gi0/0/0.40'; #Biz-internet sub-int
					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Gi0/0/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.20'; #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Gi0/0/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}if($anchor{'transport'} == 3){
					$anchor{'tloc_int'}     = 'Gi0/0/0.40'; #Biz-internet sub-int
					$anchor{'tloc_int3'}     = 'Gi0/0/0.20' if(($anchor{'r1_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int3'}     = 'Gi0/0/0.30' if(($anchor{'r1_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int3'}     = 'Gi0/0/0.50' if(($anchor{'r1_provider'} eq 'lumen')); #Private2 sub-int

					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Gi0/0/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.20' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.20')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.30' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.30')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.50' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.50')); #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Gi0/0/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}
			}else{
				if($anchor{'transport'} == 2){
				$anchor{'tloc_int'}     = 'Gi0/0/0.40';
				$anchor{'tloc_int2'}     = 'Gi0/0/0.20';
				}
				elsif($anchor{'transport'} == 3){
				$anchor{'tloc_int'}     = 'Gi0/0/0.40';
				$anchor{'tloc_int2'}     = $anchor{'tloc_int'};
				}
			}
		}
		elsif( $anchor{'router_seltype'} eq 'C8300-1N1S' ){
			$mpls_upl_int = 5;
			$anchor{'mpls_upl_int'}     = "Gi0/0/$mpls_upl_int";
			$anchor{'att_upl_int_dns'} = "g0-0-$mpls_upl_int";
			$anchor{'mci_upl_int'}     = "Gi0/0/$mpls_upl_int";
			$anchor{'isp_upl_int'}     = 'Gi0/0/4';
			$anchor{'pub_upl_int'}     = 'Gi0/0/4';
			$anchor{'cis_mls_int'}     = 'g0-0-0';
			$anchor{'cis_mls_int2'}     = 'g0-0-1';

			#Below will cover 4-transport and Dual-DIA align to the next-Gen Arch design.
			if( $anchor{'tloc'} eq 'yes' ){
				if($anchor{'transport'} == 2){
					$anchor{'tloc_int'}     = 'Gi0/0/2.40'; #Biz-internet sub-int
					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Gi0/0/2.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.20'; #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Gi0/0/2.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}if($anchor{'transport'} == 3){
					$anchor{'tloc_int'}     = 'Gi0/0/2.40'; #Biz-internet sub-int
					$anchor{'tloc_int3'}     = 'Gi0/0/2.20' if(($anchor{'r1_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int3'}     = 'Gi0/0/2.30' if(($anchor{'r1_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int3'}     = 'Gi0/0/2.50' if(($anchor{'r1_provider'} eq 'lumen')); #Private2 sub-int

					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Gi0/0/2.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.20' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.20')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.30' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.30')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.50' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.50')); #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Gi0/0/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Gi0/0/2.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}
			}else{
				if($anchor{'transport'} == 2){
				$anchor{'tloc_int'}     = 'Gi0/0/2.40';
				$anchor{'tloc_int2'}     = 'Gi0/0/2.20';
				}
				elsif($anchor{'transport'} == 3){
				$anchor{'tloc_int'}     = 'Gi0/0/2.40';
				$anchor{'tloc_int2'}     = $anchor{'tloc_int'};
				}
			}
		}
		elsif( $anchor{'router_seltype'} eq 'C8300-2N2S' ){
			$mpls_upl_int = 5;
			$anchor{'mpls_upl_int'}     = "Te0/0/$mpls_upl_int";
			$anchor{'att_upl_int_dns'} = "Te0-0-$mpls_upl_int";
			$anchor{'mci_upl_int'}     = "Te0/0/$mpls_upl_int";
			$anchor{'isp_upl_int'}     = 'Te0/0/4';
			$anchor{'pub_upl_int'}     = 'Te0/0/4';
			$anchor{'cis_mls_int'}     = 'g0-0-0';
			$anchor{'cis_mls_int2'}     = 'g0-0-1';

			#Below will cover 4-transport and Dual-DIA align to the next-Gen Arch design.
			if( $anchor{'tloc'} eq 'yes' ){
				if($anchor{'transport'} == 2){
					$anchor{'tloc_int'}     = 'Te0/1/0.40'; #Biz-internet sub-int
					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20'; #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}if($anchor{'transport'} == 3){
					$anchor{'tloc_int'}     = 'Te0/1/0.40'; #Biz-internet sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.20' if(($anchor{'r1_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.30' if(($anchor{'r1_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.50' if(($anchor{'r1_provider'} eq 'lumen')); #Private2 sub-int

					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.20')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.30')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.50')); #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}if($anchor{'transport'} == 4){
					$anchor{'tloc_int'}     = 'Te0/1/0.40'; #Biz-internet sub-int
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.20' if(($anchor{'r1_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.30' if(($anchor{'r1_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.50' if(($anchor{'r1_provider'} eq 'lumen')); #Private2 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
				}
			}else{
				if($anchor{'transport'} == 2){
				$anchor{'tloc_int'}     = 'Te0/1/0.40';
				$anchor{'tloc_int2'}     = 'Te0/1/0.20';
				}
				elsif($anchor{'transport'} == 3){
				$anchor{'tloc_int'}     = 'Te0/1/0.40';
				$anchor{'tloc_int2'}     = $anchor{'tloc_int'};
				}
			}
		}
		elsif( $anchor{'router_seltype'} eq 'C8500-12X4QC' ){
			$mpls_upl_int = 5;
			$anchor{'mpls_upl_int'}     = "Te0/0/$mpls_upl_int";
			$anchor{'att_upl_int_dns'} = "Te0-0-$mpls_upl_int";
			$anchor{'mci_upl_int'}     = "Te0/0/$mpls_upl_int";
			$anchor{'isp_upl_int'}     = 'Te0/0/4';
			$anchor{'pub_upl_int'}     = 'Te0/0/4';
			$anchor{'cis_mls_int'}     = 'Te0-0-0';
			$anchor{'cis_mls_int2'}     = 'Te0-0-1';

			#Below will cover 4-transport and Dual-DIA align to the next-Gen Arch design.
			if( $anchor{'tloc'} eq 'yes' ){
				if($anchor{'transport'} == 2){
					$anchor{'tloc_int'}     = 'Te0/1/0.40'; #Biz-internet sub-int
					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20'; #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}if($anchor{'transport'} == 3){
					$anchor{'tloc_int'}     = 'Te0/1/0.40'; #Biz-internet sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.20' if(($anchor{'r1_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.30' if(($anchor{'r1_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.50' if(($anchor{'r1_provider'} eq 'lumen')); #Private2 sub-int

					if( $anchor{'pub_provider'} =~ m/Internet/i ){
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.20')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.30')); #For dummy mpls sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'tloc_int3'} ne 'Gi0/0/0.50')); #For dummy mpls sub-int
					}else{
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
					}
				}if($anchor{'transport'} == 4){
					$anchor{'tloc_int'}     = 'Te0/1/0.40'; #Biz-internet sub-int
					$anchor{'tloc_int2'}     = 'Te0/1/0.60'; #Public-internet sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.20' if(($anchor{'r1_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.30' if(($anchor{'r1_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int3'}     = 'Te0/1/0.50' if(($anchor{'r1_provider'} eq 'lumen')); #Private2 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.20' if(($anchor{'r2_provider'} eq 'att')); #MPLS sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.30' if(($anchor{'r2_provider'} eq 'vz' )); #Private1 sub-int
					$anchor{'tloc_int4'}     = 'Te0/1/0.50' if(($anchor{'r2_provider'} eq 'lumen')); #Private2 sub-int
				}
			}
		}elsif( $anchor{'router_seltype'} eq '4461' ){
			$mpls_upl_int = 0;
			$anchor{'mpls_upl_int'}     = "Gi0/0/1";
			$anchor{'att_upl_int_dns'} = "g0-0-$mpls_upl_int";
			$anchor{'mci_upl_int'}     = "Gi0/0/1";
			$anchor{'isp_upl_int'}     = 'Gi0/0/4';
			$anchor{'cis_mls_int'}     = 'g0-0-1';
			$anchor{'cis_mls_int2'}     = 'g0-0-3';

			#Below will cover 4-transport and Dual-DIA align to the next-Gen Arch design.
			if($anchor{'transport'} == 2){
			$anchor{'tloc_int'}     = 'Te0/0/4.40';
			$anchor{'tloc_int2'}     = 'Te0/0/4.20';
			}
			elsif($anchor{'transport'} == 3){
			$anchor{'tloc_int'}     = 'Te0/0/4.40';
			$anchor{'tloc_int2'}     = $anchor{'tloc_int'};
			}
		}
		else{
			$mpls_upl_int = 0;
			$anchor{'mpls_upl_int'}     = "Gi0/0/2";
			$anchor{'att_upl_int_dns'} = "g0-0-$mpls_upl_int";
			$anchor{'mci_upl_int'}     = "Gi0/0/$mpls_upl_int";
			$anchor{'isp_upl_int'}     = 'Gi0/0/0';
			$anchor{'cis_mls_int'}     = 'g0-0-1';
			$anchor{'cis_mls_int2'}     = 'g0-0-3';

			if(( $anchor{'router_seltype'} =~ /^(4331|4451|ASR)$/ ) and ($anchor{'transport'} == 2)){
			$anchor{'tloc_int'}     = "Gi0/0/2.40";
			$anchor{'tloc_int2'}     = "Gi0/0/2.20";
			}
			elsif(( $anchor{'router_seltype'} =~ /^(4331|4451|ASR)$/ ) and ($anchor{'transport'} == 3)){
			$anchor{'tloc_int'}     = "Gi0/1/0.40";
			$anchor{'tloc_int2'}     = "Gi0/0/2.40";
			}
		}
	}

	#Anchor values for SDWAN Next-Gen csv templates 2-transport
	my $tloc_int = $anchor{'tloc_int'};
	if($tloc_int =~ m/Gi/i){
		my($biz_tloc_int) = substr($tloc_int, 2,8);
		$anchor{'biz_tloc'} = "GigabitEthernet$biz_tloc_int";
	}if($tloc_int =~ m/Te/i){
		my($biz_tloc_int) = substr($tloc_int, 2,8);
		$anchor{'biz_tloc'} = "TenGigabitEthernet$biz_tloc_int";
	}

	#Anchor values for SDWAN Next-Gen csv templates 2-transport
	my $tloc_int2 = $anchor{'tloc_int2'};
	if($tloc_int2 =~ m/Gi/i){
		my($pub_tloc_int2) = substr($tloc_int2, 2,8);
		$anchor{'pub_tloc'} = "GigabitEthernet$pub_tloc_int2";
	}if($tloc_int2 =~ m/Te/i){
		my($pub_tloc_int2) = substr($tloc_int2, 2,8);
		$anchor{'pub_tloc'} = "TenGigabitEthernet$pub_tloc_int2";
	}

	#Anchor values for SDWAN csv templates
	my $tloc_int3 = $anchor{'tloc_int3'};
	if($tloc_int3 =~ m/Gi/i){
		my($mpls_tloc_int) = substr($tloc_int3, 2,8);
		$anchor{'mpls_tloc'} = "GigabitEthernet$mpls_tloc_int";
	}if($tloc_int3 =~ m/Te/i){
		my($mpls_tloc_int) = substr($tloc_int3, 2,8);
		$anchor{'mpls_tloc'} = "TenGigabitEthernet$mpls_tloc_int";
	}

	#Anchor values for SDWAN Next-Gen csv templates 3 and 4 -transport with mpls ckt on R2
	my $tloc_int4 = $anchor{'tloc_int4'};
	if($tloc_int4 =~ m/Gi/i){
		my($mpls_tloc_int2) = substr($tloc_int4, 2,8);
		$anchor{'mpls_tloc2'} = "GigabitEthernet$mpls_tloc_int2";
	}if($tloc_int4 =~ m/Te/i){
		my($mpls_tloc_int2) = substr($tloc_int4, 2,8);
		$anchor{'mpls_tloc2'} = "TenGigabitEthernet$mpls_tloc_int2";
	}

	if ( $anchor{'router_type'} =~ /^(2951|3945|3945E)$/ ) {
		$anchor{'wan_int_netflow'} = smbReadFile("Modules/isr_int_netflow.txt");
	} else {
		$anchor{'wan_int_netflow'} = smbReadFile("Modules/asr_int_netflow.txt");
	}
	#### IPS variables sets
	if ( $anchor{'ips'} eq 'Y' ) {
		$anchor{'ips1_tad'} = 'IPS1';
		$anchor{'ips2_tad'} = 'IPS2';
	} else {
		$anchor{'ips1_tad'} = 'empty';
		$anchor{'ips2_tad'} = 'empty';
	}

	# ACL 59, 60 and 172 variables sets
	# Added ACL 59 specific for 9800 line 180 for DNAC
	$anchor{'acl_59'}  = smbReadFile("Modules/acl_59.txt");
	$anchor{'acl_59_9800'}  = smbReadFile("Modules/acl_59_9800.txt"); ### This will be part of the clean-UP both acl_59 and this have the same values now
	$anchor{'acl_60'}  = smbReadFile("Modules/acl_60.txt");
	$anchor{'acl_172'} = smbReadFile("Modules/acl_172.txt");

	# New module sets
	if ($anchor{'man_campus'} eq 'Sierra Nevada'){
		$anchor{'acl_eac'}          = smbReadFile("Modules/acl_eac_sierra.txt");
	}
	else{
		$anchor{'acl_eac'}          = smbReadFile("Modules/acl_eac.txt");
	}
	$anchor{'aruba_acl_eac'}    = smbReadFile("Modules/aruba_acl_eac.txt");
	$anchor{'acl_voice'}        = smbReadFile("Modules/acl_voice.txt");
	$anchor{'aruba_acl_voice'}  = smbReadFile("Modules/aruba_acl_voice.txt");
	$anchor{'banner'}           = smbReadFile("Modules/banner.txt");
	$anchor{'cis_snmp'}         = smbReadFile("Modules/cis_snmp.txt");
	$anchor{'cis_asr_snmp'}     = smbReadFile("Modules/cis_asr_snmp.txt");
	$anchor{'cis_aaa'}          = smbReadFile("Modules/cis_aaa.txt");
	$anchor{'cis_qos'}          = smbReadFile("Modules/cis_qos.txt");
	$anchor{'cis_qos_acl'}      = smbReadFile("Modules/cis_qos_acl.txt");
	$anchor{'mls_qos'}          = smbReadFile("Modules/mls_qos.txt");
	$anchor{'mls_qos_acl'}      = smbReadFile("Modules/mls_qos_acl.txt");
	$anchor{'xl_mls_snmp'}      = smbReadFile("Modules/xl_mls_snmp.txt");
	$anchor{'m_mls_snmp'}       = smbReadFile("Modules/m_mls_snmp.txt");
	$anchor{'stk_qos'}          = smbReadFile("Modules/stk_qos.txt");
	$anchor{'stk_qos_acl'}      = smbReadFile("Modules/stk_qos_acl.txt");
	$anchor{'aruba_stk_qos_acl'}      = smbReadFile("Modules/aruba_stk_qos_acl.txt");
	$anchor{'aruba_stk_snmp'}         = smbReadFile("Modules/aruba_stk_snmp.txt");
	$anchor{'stk_snmp'}         = smbReadFile("Modules/stk_snmp.txt");
	$anchor{'aruba_stk_snmp'}     = smbReadFile("Modules/aruba_stk_snmp.txt");
	$anchor{'p_stk_qos'}        = smbReadFile("Modules/p_stk_qos.txt");
	$anchor{'acl_segmentation'} = smbReadFile("Modules/acl_segmentation.txt");
	$anchor{'aruba_acl_segmentation'} = smbReadFile("Modules/aruba_acl_segmentation.txt");

	# Add BFD to BGP config for MPLS Ethernet locations
	if ( $wantype eq 'MPLS_Ethernet' ) {
		$anchor{'cis1_bgp_bfd'} = "neighbor !pri_wan_ip_per! fall-over bfd";
		$anchor{'cis2_bgp_bfd'} = "neighbor !sec_wan_ip_per! fall-over bfd";
	} else {
		$anchor{'cis1_bgp_bfd'} = '!';
		$anchor{'cis2_bgp_bfd'} = '!';
	}

	# Internet Provider
	my %internet_provider = (
						 'Internet-Comcast',    		'Comcast - 1-800-741-4141',
						 'Internet-Spectrum', 		'Spectrum - 1-888-812-2591',
						 'Internet-Century Link', 	'Century Link 1-877-453-8353',
						 'Internet-COX3',  			'COX - 1-877-715-1327',
						 'Internet-Granite', 		'Granite - 888-375-3755',
						 'Internet-ATT', 			'ATT - 1-888-613-6330',
						 'Internet-Verizon', 		'Verizon 1-866-553-1226',
						 'Internet-Lumen', 			'Lumen - 1-877-453-8353',
						 'Internet-Globe PH', 			'Globe - +63-9176883262',
						 'Internet-PLDT PH', 			'PLDT - pldtecsait-bpm@pldt.com.ph',
						 'Internet-Converge PH', 			'Converge - +63-2-8667-0800',
						 'Internet-ETPI PH', 			'ETPI - +63-2-5300-7000',
						 'Internet-Crown Castle', 			'Crown Castle - 1-855-933-4237',
						 'Internet-Frontier', 			'Frontier - 1-888-637-9620',
						 'Internet-LightPath', 			'LightPath - 1-877-544-4872',
						 'Internet-Ziply Fiber', 			'Ziply Fiber - 1-888-488-0072',
	);
	my $int_p = $anchor{'int_provider'};
	$anchor{'isp'} = uc($int_p);
	my $pub_p = $anchor{'pub_provider'};
	$anchor{'isp2'} = uc($pub_p);
	$anchor{'isp_vendor'} = substr($anchor{'isp'}, 9);
	$anchor{'isp2_vendor'} = substr($anchor{'isp2'}, 9);
	$anchor{'int_provider_contact'} = $internet_provider{$int_p};
	$anchor{'pub_provider_contact'} = $internet_provider{$pub_p};

	#removing "Ethernet" under MPLS_Ethernet for interface description
	if ( $anchor{'pri_circuit_type'} =~ /^(MPLS_Ethernet)$/ ) {
	$anchor{'mpls'} = substr($anchor{'pri_circuit_type'}, 0,4);
	$anchor{'mpls2'} = substr($anchor{'sec_circuit_type'}, 0,4);
	}

	# MPLS Provider Color and Carrier
	my %mpls_description = (
						 'att',   'ATT-UHG-MPLS',
						 'verizon', 'VZ-UHG-MPLS',
						 'vz', 'VZ-UHG-MPLS',
						 'lumen', 	'LUMEN-UHG-MPLS',
	);


	my %mpls_provider = (
						 'att', => {'color' => 'mpls', 'carrier' => 'carrier1','bgp' => 'UHGATTbgpyI6HYRM','asn' => '13979','contact' => '1-800-732-5980'},
						 'verizon' => {'color' => 'private1', 'carrier' => 'carrier2','bgp' => 'UHGVERbgpbqZyofa','asn' => '65000','contact' => '1-866-553-1226'},
						 'vz' => {'color' => 'private1', 'carrier' => 'carrier2','bgp' => 'UHGVERbgpbqZyofa','asn' => '65000','contact' => '1-866-553-1226'},
						 'lumen' => {'color' => 'private2', 'carrier' => 'carrier5','bgp' => 'UHGCTLbgpbRjDiGs','asn' => '209','contact' => '1-877-453-8353'}
	);


	my $mpls_p = $anchor{'r1_provider'};
	my $mpls_p2 = $anchor{'r2_provider'};
	$anchor{'r1_provider'} = uc($anchor{'r1_provider'});
	$anchor{'r2_provider'} = uc($anchor{'r2_provider'});
	$anchor{'pri_provider'} = uc($anchor{'r1_provider'});
	$anchor{'sec_provider'} = uc($anchor{'r2_provider'});
	$anchor{'r1_provider'} = 'VZ' if ($anchor{'r1_provider'} eq 'VERIZON');
	$anchor{'r2_provider'} = 'VZ' if ($anchor{'r2_provider'} eq 'VERIZON');
	$anchor{'r1_mpls_description'} = $mpls_description{$mpls_p};
	$anchor{'r2_mpls_description'} = $mpls_description{$mpls_p2};


	$anchor{'pri_provider_contact'} = $mpls_provider{$mpls_p}{'contact'};
	$anchor{'sec_provider_contact'} = $mpls_provider{$mpls_p2}{'contact'};
	$anchor{'r1_color'} = $mpls_provider{$mpls_p}{'color'};
	$anchor{'r1_color_uc'} = uc($anchor{'r1_color'});
	$anchor{'r1_carrier'} = $mpls_provider{$mpls_p}{'carrier'};
	$anchor{'r2_color'} = $mpls_provider{$mpls_p2}{'color'};
	$anchor{'r2_color_uc'} = uc($anchor{'r2_color'});
	$anchor{'r2_carrier'} = $mpls_provider{$mpls_p2}{'carrier'};
	$anchor{'r1_bgp_password'} = $mpls_provider{$mpls_p}{'bgp'};
	$anchor{'r1_bgp_neighbor'} = $mpls_provider{$mpls_p}{'asn'};
	$anchor{'r2_bgp_password'} = $mpls_provider{$mpls_p2}{'bgp'};
	$anchor{'r2_bgp_neighbor'} = $mpls_provider{$mpls_p2}{'asn'};

	if ($anchor{'pri_vlan'} > 0){
		$anchor{'r1_vlan'} = '.' . $anchor{'pri_vlan'};
		$anchor{'mtu1'} = '1504';
	}
	else{
		$anchor{'r1_vlan'} = '';
		$anchor{'mtu1'} = '1500';
	}

	if ($anchor{'r2vlan'} > 0){
		$anchor{'r2_vlan'} = '.' . $anchor{'r2vlan'};
		$anchor{'mtu2'} = '1504';
	}
	else{
		$anchor{'r2_vlan'} = '';
		$anchor{'mtu2'} = '1500';
	}

	# Stack model
	$anchor{'stk_model'} = '6300M';
	$anchor{'stk_model'}  = 'C9300-48P-E' if ( $anchor{'stack_vendor'} eq 'cisco' );

	#Additional Radius Dynamic Authors / AAA server group for PH sites
	if ( $anchor{'site_code'} =~ /^PH/ and $anchor{'stack_vendor'} eq 'cisco'){
		$anchor{'dynamic_author_addtl'} = smbReadFile("Modules/dyn_author_ph.txt");
	}else{
		$anchor{'dynamic_author_addtl'} = '!';
	}

	#Internet subnetmask
	my $full_mask = $anchor{'int_cer_ip'};
	my($network, $netbit) = split /\//, $full_mask;
	my $mask  = (2 ** $netbit - 1) << (32 - $netbit);
	$anchor{'int_cer_mask'} = join( '.', unpack( "C4", pack( "N", $mask ) ) );
	$anchor{'int_cer_net'} = $network;

	#Public-internet subnetmask
	my $full_mask_pub = $anchor{'pub_cer_ip'};
	my($network_pub, $netbit_pub) = split /\//, $full_mask_pub;
	my $mask_pub  = (2 ** $netbit_pub - 1) << (32 - $netbit_pub);
	$anchor{'pub_cer_mask'} = join( '.', unpack( "C4", pack( "N", $mask_pub ) ) );
	$anchor{'pub_cer_net'} = $network_pub;
}

sub inputErrorCheckGlobal {

prtout("$anchor{'int_bw'}");

my $ct = '0';

	if ($anchor{'loop_oct1'} == '0'){
			prtout( "" );
			prtout( "ERROR: Please input a value in the 1st octet subnet: eg: 10 or 172." );
			$ct++;
	}

	if ($anchor{'int_bw'} !~ m/\//){
			prtout( "" );
			prtout( "ERROR: Please input the proper download and upload for the internet bandwidth field. eg: 10000/10000." );
			$ct++;
	}

	if (     $anchor{'pri_circuit_type'} =~ /(?:DS3_subrate|Metro_Ethernet|MPLS_Ethernet)/
		 and $anchor{'subrate'} * 1000 < 1 )
	{
		prtout( "Either the subrate/ethernet port speed was not entered or it is incorrect.",
				"Please enter this information on the Configurator input page and rerun." );
		$ct++;
	}
	if ( $anchor{'ips'} eq 'Y' ) {
		prtout(
			"Currently, Configurator does not support the IPS configuration. ",
"Please select N for the IPS field on the Configurator input page and add the configuration per current standards documentation."
		);
		$ct++;
	}
	if ( $anchor{'mci_gold_car'} < 1 ) {
		prtout(
"You have no entered a valid value for the Verizon Gold CAR. Please enter this information on the Configurator input page and rerun."
		);
		$ct++;
	}
	if ( $anchor{'uhgdivision'} eq '' ) {
		prtout(
"You have not entered a valid value for the UHG Division. Please enter this information on the Configurator input page and rerun."
		);
		$ct++;
	}

	if ( $anchor{'wlan'} eq 'Y' and $anchor{'uhgdivision'} eq 'TRICARE' ) {
		prtout( "Currently, Configurator does not support a wireless configuration for a TriCare site." );
		#xit(1);
		$ct++;
	}
	if ( $anchor{'region'} eq 'USA' and length( $anchor{'state'} ) > 2 ) {
		prtout(
			"State: '$anchor{'state'}' (quoted for whitespace visibility)",
"It appears the USA State has more than two characters. Please check the NSLANWAN site to ensure your site is using a correct two letter state abbreviation.",
"Please also ensure that the state abbreviation entered on the NSLANWAN website does not contain any spaces at the end of the abbreviation.",
			''
		);
		$ct++;
	}

	#script will terminate
	if ($ct > 0){
		xit(1);
	}
}

# Replaces the smbclient 'get' since it stopped working suddenly
sub smbGet {
	( my $remoteFile, my $localFile ) = @_;

	# Printing out dots as a progress indicator since the browser is timing out suddenly. 5/25/2017

	my @rtab = $smb->stat("smb://$SMB_SHARE_PATH/$remoteFile");
	my @ltab = stat($localFile);
	# compare the file size and modified date
	if ( ( $rtab[7] == $ltab[7] ) && ( $rtab[10] == $ltab[10] ) ) {
	} else {
		my $size = 0;
		my $fd = $smb->open("smb://$SMB_SHARE_PATH/$remoteFile") or do {
			print "ERROR: open template file '$SMB_SHARE_PATH/$remoteFile' failed with error:\n";
			print "$!\n";
			xit(1);
		};
		open( LOCF, '>', $localFile ) or do {
			print "ERROR: open local file '$localFile' failed with error:\n";
			print "$!\n";
			xit(1);
		};
		while ( my $line = $smb->read( $fd, 1024 ) ) {
			$size += length($line);
			if ( $size > 5000000 ) {
				$size = 0;
				print ".";    # zzz prevent timeouts during execution
			}
			print LOCF $line;
		}
		$smb->close($fd);
		close(LOCF);

	}

}

# Replaces the smbclient 'put' since it stopped working suddently
sub smbPut {
	( my $localFile, my $remoteFile ) = @_;
	my $fd = $smb->open( ">smb://$SMB_SHARE_PATH/$remoteFile", 0666 ) or do {
		print "ERROR: create file '$SMB_SHARE_PATH/$remoteFile' failed: $!\n";
		xit(1);
	};
	my $buf;
	open( LOCF, $localFile ) or do {
		print "open local file '$localFile' failed: $!\n";
		xit(1);
	};
	while ( read( LOCF, $buf, 1024 ) ) {
		$smb->write( $fd, $buf ) or do {
			print "ERROR: write to file '$remoteFile' failed: $!\n";
			xit(1);
		};
	}
	close(LOCF);
	$smb->close($fd);
}

# Read data via SMB module
# This uses a ref to a copy of the anchor hash so iterations can resolve to different values
#
# Change starting with version 16.0.1 - Copy file locally, then read it. This gets away from SMB reads
sub smbReadFile {
	my $file = shift;
	my $href = shift
	  || \%anchor;    # allows passing a copy of the hash so values can be updated dynamically
	my $data;         # 
	my @data;

	# SMB copy file locally and read it
	my $localFileTemp = $ROOTDIR . '/' . basename($file) . ".$$";    # Append PID for uniqueness
	my $srcFile       = "$SMB_TEMPLATE_DIR/$file";

	# Turning this into a command so I can print it. SMB copies are failing for some reason. :/
	smbGet( "$srcFile", $localFileTemp );

	if ( !( -e $localFileTemp ) ) {
		print "Copy template file '$srcFile' to '$localFileTemp' failed\n";
		print "Please contact development and provide this error.\n";
		xit(1);
	}

	# Read the file locally - skip empty lines
	open( TEMPLATE, $localFileTemp ) or do {
		print "Read file '$srcFile' failed: $!\n";
		print "Please contact development and provide this error\n";
		unlink($localFileTemp);    # cleanup on exit
		xit(1);
	};

	# Read the file, strip all CR/LF chars and save to array for parsing
	while (<TEMPLATE>) {
		push @data, $_;
	}
	close(TEMPLATE);
	unlink($localFileTemp);        # we're done with this template so remove it

	# iterate the array containing the file contents and replace meta values with actual values
	# Since we will 'edit' each line in place, keep track of the array index while looping
	for ( my $ct = 0 ; $ct <= $#data ; $ct++ ) {
		my $line   = $data[$ct];    # this line will be parsed,
		my $newval = '';            # and replaced with this value when it is built
		while ( $line =~ /\!([a-z_0-9]+)\!/ ) {    # found an anchor symbol
			my $anchorsymbol = $1;                 # this is our symbol
			$newval .= $`;                         # everything before the symbol can go into the new value
			my $postmatch = $';                    # everything after the symbol - it may contain more symbols

			# Replace the symbol with a value if one exists
			if ( defined $$href{$anchorsymbol} ) {

				# If the replacement value contains symbols it will need to be processed again to resolve them
				if ( $$href{$anchorsymbol} =~ /\![a-z_0-9]+\!/ ) {
					$line = $$href{$anchorsymbol} . $postmatch;    # stuff the remaining portion back into $line
					                                               # and parse it again

					# Replacement value has no anchors - save it to our new value
				} else {
					$newval .= $$href{$anchorsymbol};
					$line = $';
				}

				# There is no value for the anchor symbol - leave it as a literal
			} else {
				$newval .= '!' . $anchorsymbol . '!';
				$line = $';
			}
		}
		$newval .= $line
		  if ( $line ne '' );    # if the old line has anything left in it, add it to our new value
		$data[$ct] = $newval;    # overwrite the old value with the new one
	}
	my $recsep = sprintf( "0x%X", 0 );
	$data = join( $recsep, @data );    # join with null chars - don't touch CR/LF since record separators may differ
	return $data;
}

# Reads in template via the smbReadFile subroutine, which replaces !symbols! with values
# Then writes out the file on the destination host.
sub writeTemplate {
	( my $inputFile, my $outputFile ) = @_;

	# Resolve anchor value representations prior to writing
	# Use a copy of the anchor hash so iterations can use different values
	my %anchor_tmp = %anchor;

	# Replace all symbols with values in the copy of the anchor hash
	my $err = resolveAnchorValues( \%anchor_tmp );

	# split on null characters
	my $recsep = sprintf( "0x%X", 0 );
	my @data = split( /$recsep/, smbReadFile( $inputFile, \%anchor_tmp ) );

	# Make this sucker whether you need to or not
	mkdir("$SVR_ROOTDIR/$OutputDir");

	# write out the file - zzz do this locally and copy in the future
	# Exit if any anchor values were not replaced in the file
	my %unresolvedAnchor;
	my $lastline;

	open( OF, "+>$SVR_ROOTDIR/$OutputDir/$outputFile" );

	foreach (@data) {
		my $line = $_;
		$lastline = $line;
		print OF $line;

		#	$smb->write( $smbOut, $line );
		if ( $line =~ /\!([a-z_0-9]+)\!/ ) {
			$unresolvedAnchor{$1} = '';
		}
	}

	# Add newline at EOF if one is needed
	if ( $lastline !~ /[\n]$/ ) {
		print OF "\n";
	}
	close OF;

	# Exit if there are unresolved anchor symbols
	if ( scalar( keys %unresolvedAnchor ) ) {
		print "      ERROR: The following file contains unresolved anchor symbols:\n";
		print "   Template: $inputFile\n";
		print "Output file: $outputFile\n";
		foreach my $key ( sort keys %unresolvedAnchor ) {
			print "     Anchor: $key\n";
		}
	}
}

# set the stacks for the site
sub getstacks {
	( my $id, my $sitecode ) = @_;
	my @stack = ();
	my $query = $dbh->prepare(
"SELECT concat('stk', mailroute, bldg, if(floor < 10, '0', ''), floor, idf, '0', stack, ',', switchamt) FROM configurator_stacks WHERE config_id='$id' ORDER BY stackid"
	  )
	  or do {
		prtout( "Query '$query' failed: " . $dbh->errstr );
		xit(1);
	  };
	$query->execute;
	while ( ( my $stack ) = $query->fetchrow_array ) {
		push @stack, $stack;
	}
	return @stack;
}

# Resolve anchor symbols with actual values
# Exit if an anchor symbol does not have a corresponding value
sub resolveAnchorValues {
	my $href = shift;

	# Iterate through every anchor key/value pair and replace symbols with actual values
	foreach my $anchor ( sort keys %$href ) {
		my $orig_value = $$href{$anchor};
		next if ( !$orig_value );
		my $new_value = '';
		while ( $orig_value =~ /\!([a-z_0-9]+)\!/ ) {
			my $anchor_symbol = $1;
			my $postmatch     = $';    # need to save the post-match value because it's needed after the regex below
			$new_value .= $`;

			# Replace the symbol with a value if one exists
			if ( defined $$href{$anchor_symbol} ) {

				# If the replacement value contains symbols it will need to be processed again to resolve them
				if ( $$href{$anchor_symbol} =~ /\![a-z_0-9]+\!/ ) {
					$orig_value = $$href{$anchor_symbol} . $postmatch;

					# Replacement value has no anchors - save it to our new value
				} else {
					$new_value .= $$href{$anchor_symbol};
					$orig_value = $';
				}

				# There is no value for the anchor symbol
			} else {
				$new_value .= '!' . $anchor_symbol . '!';
				$orig_value = $';
			}
		}

		# If there's anything left in the original value, add to the new value
		$new_value .= $orig_value if ( $orig_value ne '' );

		# Now this key/value pair should be fully resolved
		$$href{$anchor} = $new_value;
	}
}

# Modify the ID of the Visio page
sub VisioReIDTab {    # was VisioReidTab
	( my $reidto, my $data ) = @_;
	my $pgstart = index( $$data, "'" );
	my $pgend = index( $$data, "'", $pgstart + 1 );
	substr( $$data, $pgstart, $pgend - $pgstart + 1 ) = "'$reidto'";
}

# pass ref to scalar containing Visio data
# Returns a hash of page ids containing strings where the pages start
sub VisioReadTabs {
	my $data = shift;
	my %pgcache;
	my ( $pgstart, $pgid, $tagend, $pgend );
	prtout("Finding tabs in the Visio Templates");
	$pgstart = index( $$data, "<Page ID" );
	while ( $pgstart != -1 ) {
		$pgid = substr( $$data, $pgstart + 9, 4 );
		$pgid =~ s/\s//g;
		$pgid =~ s/\'//g;
		$pgid   = scalar($pgid);
		$tagend = index( $$data, '>', $pgstart );
		$pgend  = index( $$data, '</Page>', $pgstart );
		$pgcache{$pgid} = substr( $$data, $pgstart, $pgend - $pgstart + 7 );
		substr( $$data, $pgstart, $pgend - $pgstart + 7 ) = '';
		$pgstart = index( $$data, "<Page ID" );    # used to find EOF
	}
	return %pgcache;
}

sub VisioRenameTab {
	( my $renameto, my $data ) = @_;
	my $pagepos = index( $$data, '<Page ID=' );
	my $nameupos = index( $$data, "NameU='", $pagepos );
	$nameupos += 7;
	my $nameuend = index( $$data, "'", $nameupos );
	substr( $$data, $nameupos, $nameuend - $nameupos ) = $renameto;
	my $namepos = index( $$data, "Name='", $nameupos );
	$namepos += 6;
	my $nameend = index( $$data, "'", $namepos );
	substr( $$data, $namepos, $nameend - $namepos ) = $renameto;
}

sub VisioControlLayer {
	( my $findlayer, my $visible, my $data ) = @_;
	my $layerpos   = index( $$data, "<Name>$findlayer</Name>" );
	my $visiblepos = index( $$data, "<Visible>", $layerpos );
	my $printpos   = index( $$data, "<Print>", $visiblepos );
	if ( $layerpos ne '-1' ) {
		substr( $$data, $visiblepos + 9, 1 ) = $visible;
		substr( $$data, $printpos + 7,   1 ) = $visible;
	} else {
		my $callingsub  = ( caller 0 )[3];
		my $callingline = ( caller 0 )[2];
		debug( "Layer '$findlayer' not found in the Visio", "Caller: $callingsub line: $callingline" );
		print "Layer '$findlayer' is not in the Visio document - skipping\n";
	}
}

sub writeRemoteSiteBuildChecklist {
return unless ($anchor{'proj_type'} eq 'build');

my $outputFile = writeTemplate( "checklist/Remote Site Build Checklist.xml", $anchor{'site_code'} . ' - Remote Site Build Checklist.xls' );



	return $outputFile;
}

sub deviceType {
	my $name = shift;
	if ( $name =~ /^cis/ ) {
		return 'Router';
	} elsif ( $name =~ /^(?:mls|stk)/ ) {
		return 'Data Switch';
	} else {
		return 'Network Appliance';
	}
}

sub writeRow {
	( my $ws, my $row, my $aref, my $fmt ) = @_;
	if ( defined $fmt ) {
		$ws->write_row( $row, 0, $aref, $fmt );
	} else {
		$ws->write_row( $row, 0, $aref );
	}
	$row++;
	return $row;
}

# Insert script line numbers into print output if we're debugging
sub prtout {
	foreach (@_) {
		print $_ . "\n";
	}
}

sub xit() {
	my $xitval = shift || 0;
	exit $xitval;
}

sub devStrLen {
	my $item   = shift;
	my $string = shift;
	if ( defined $string ) {
		print "$item len is: " . length($string) . "\n";
	} else {
		print "$item\n";
	}
}

sub debug {
	return unless ($DEBUG);
	foreach (@_) {
		print "DEBUG: $_\n";
	}
}

sub getTime {
#define Date and Time
	my @months = qw( Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec );
	my @days = qw(Sun Mon Tue Wed Thu Fri Sat Sun);
	my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime();
	$year = $year + 1900;
	my $dtime = ("$months[$mon] $mday $year, $hour:$min");
	return $dtime;
}

sub compress {
	die "SVR_ROOTDIR is not defined" unless $SVR_ROOTDIR;
	die "OutputDir is not defined" unless $OutputDir;

	my $directory = "$SVR_ROOTDIR/$OutputDir";
	my $zipfile   = "$SVR_ROOTDIR/$OutputDir.zip";

	# List of folders to create
	my @folders = qw(Archive Misc Configs Orders Wireless);
	my $pictures_folder = "$directory/Pictures";
	my $projectname_folder = "$pictures_folder/ProjectName";

	# Create folders and add .keep placeholder
	foreach my $folder (@folders) {
		my $folder_path = "$directory/$folder";
		make_path($folder_path);
	}

	# Create Pictures and nested ProjectName folder
	make_path($projectname_folder);
	#print "Created nested folder: $projectname_folder\n";

	# Move files based on extension or name pattern
	opendir(my $dh, $directory) or die "Cannot open directory $directory: $!";
	while (my $file = readdir($dh)) {
		next if $file =~ /^\./;  # Skip hidden files and . ..
		my $full_path = "$directory/$file";
		next unless -f $full_path;  # Skip if not a file

		if ($file =~ /\.(txt|csv)$/i) {
			move($full_path, "$directory/Configs/$file") or warn "Failed to move $file to Configs";
		}
		elsif ($file =~ /\.doc$/i || $file =~ /Checklist\.xls$/i) {
			move($full_path, "$directory/Misc/$file") or warn "Failed to move $file to Misc";
		}
	}
	closedir($dh);

	print "Compressing contents of [$directory] into zipfile [$zipfile]\n";

	# Change to $directory to make paths relative
	my $cwd = getcwd();
	chdir($directory) or die "Cannot change to directory $directory: $!";

	# Collect all files excluding folders
	my @files;
	find(
		sub {
			return if -d;

			# Normalize path
			my $rel_path = $File::Find::name;
			$rel_path =~ s{^\./}{};

			# Skip files inside folders
			return if $rel_path =~ m{^(Archive|Misc|Configs|Orders|Wireless|Pictures/)};

			push @files, $File::Find::name;
		},
		'.'
	);

	# Step 1: Zip flattened files (only if any)
	if (@files) {
		my $quoted_files = join(' ', map {
			my $f = $_;
			$f =~ s/'/'\\''/g;
			"'$f'"
		} @files);

		my $flat_cmd = "zip -j '$zipfile' $quoted_files > /dev/null 2>&1";
		system($flat_cmd) == 0 or die "Zip command (flat files) failed: $?";
	} else {
		print "No flat files to zip.\n";
	}

	# Step 2: Add folders with structure (suppress output)
	foreach my $folder (@folders, 'Pictures') {
		my $folder_cmd = "zip -r '$zipfile' '$folder' > /dev/null 2>&1";
		system($folder_cmd) == 0 or warn "Zip command (folder $folder) failed: $?";
	}

	# Return to original working directory
	chdir($cwd);

	return $zipfile; 

}

sub stripSubInt {
	my $clean_int = shift;
	$clean_int = (split(/\./, $clean_int))[0]; # Split on the dot and take the first part

	return $clean_int;
}