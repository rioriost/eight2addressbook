#!/usr/bin/php -q
<?php
#============================================================================
# This library is free software; you can redistribute it and/or
# modify it under the terms of version 2.1 of the GNU Lesser General Public
# License as published by the Free Software Foundation.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
#============================================================================
# Copyright (C) 2017 Ryo Fujita
# Author: Ryo Fujita <rio@rio.st>
#============================================================================

# This script converts CSV file in UTF-8 downloaded from https://8card.net/export/csv_orders to VCF.
# This script requires Yahoo APP ID to use Japanese Morphological Analysis and YOLP.
# See below
# https://developer.yahoo.co.jp/webapi/jlp/ma/v1/parse.html
# https://developer.yahoo.co.jp/webapi/map/openlocalplatform/v1/addressdirectory.html

define("YAHOO_APP_ID", "");
define("MA_REQ_URL", "https://jlp.yahooapis.jp/MAService/V1/parse?appid=" . YAHOO_APP_ID . "&results=ma&sentence=%s");
define("DIR_REQ_URL", "https://map.yahooapis.jp/search/address/V1/addressDirectory?ac=%s&mode=2&appid=" . YAHOO_APP_ID);

// Template of VCF
define("VCF_TMPLT", "BEGIN:VCARD
VERSION:3.0
N:{FAMILYNAME};{GIVENNAME};;;
FN:{FULLNAME}
X-PHONETIC-FIRST-NAME:{GIVENNAMEPHN}
X-PHONETIC-LAST-NAME:{FAMILYNAMEPHN}
ORG:{ORG};{DEPT}
TITLE:{TITLE}
EMAIL;type=INTERNET;type=WORK;type=pref:{EMAIL}
TEL;type=WORK;type=VOICE;type=pref:{TEL}
item1.ADR;type=WORK;type=pref:;;{BUILDING};{ADDRESS};{PREF};{ZIP};
item1.X-ABADR:jp
URL;type=WORK;type=pref:{URL}
END:VCARD");

// Eight adds header lines to CSV
define("SKIP_LINE_NO", 8);

$csv_field_order = array(
	"ORG"=>0,
	"DEPT"=>1,
	"TITLE"=>2,
	"FULLNAME"=>3,
	"EMAIL"=>4,
	"ZIP"=>5,
	"ADDRESS"=>6,
	"TEL"=>7,
	"CELLPHONE"=>11,
	"URL"=>12,
	"REGDATE"=>13,
);

$pref_ar = array("01" => "北海道", "02" => "青森県", "03" => "岩手県", "04" => "宮城県", "05" => "秋田県", "06" => "山形県", "07" => "福島県", "08" => "茨城県", "09" => "栃木県", "10" => "群馬県", "11" => "埼玉県", "12" => "千葉県", "13" => "東京都", "14" => "神奈川県", "15" => "新潟県", "16" => "富山県", "17" => "石川県", "18" => "福井県", "19" => "山梨県", "20" => "長野県", "21" => "岐阜県", "22" => "静岡県", "23" => "愛知県", "24" => "三重県", "25" => "滋賀県", "26" => "京都府", "27" => "大阪府", "28" => "兵庫県", "29" => "奈良県", "30" => "和歌山県", "31" => "鳥取県", "32" => "島根県", "33" => "岡山県", "34" => "広島県", "35" => "山口県", "36" => "徳島県", "37" => "香川県", "38" => "愛媛県", "39" => "高知県", "40" => "福岡県", "41" => "佐賀県", "42" => "長崎県", "43" => "熊本県", "44" => "大分県", "45" => "宮崎県", "46" => "鹿児島県", "47" => "沖縄県");

$dir_req_cache = array();

// Show Usage and exit
function usage(){
	$com_name = basename($_SERVER["PHP_SELF"]);
	echo
"
Usage: " . $com_name . " [OPTIONS] InputFile OutputDirectory

Options:
  --help, -h		Show this usage guide
  --verbose, -v		Show contextual information and format for easy reading
"	
	;
	exit(0);
}

// Morphological Analysis to devide Full name into Family and Given names.
function get_morphological_analysis($sentence){
	$buf = file(sprintf(MA_REQ_URL, urlencode($sentence)));
	$xml = simplexml_load_string(implode("", $buf));
	$json = json_encode($xml);
	$array = json_decode($json, TRUE);
	$ret = array();
	switch((int)$array["ma_result"]["filtered_count"]){
	case 2: // Devided succesfully
		foreach($array["ma_result"]["word_list"]["word"] as $word){
			if($word["pos"] == "名詞"){
				if(isset($ret["FAMILYNAME"]) == false){
					$ret["FAMILYNAME"] = $word["surface"];
					$ret["FAMILYNAMEPHN"] = $word["reading"];
				} else {
					$ret["GIVENNAME"] = $word["surface"];
					$ret["GIVENNAMEPHN"] = $word["reading"];
				}
			}
		}
		break;
	case 1:
		$word = $array["ma_result"]["word_list"]["word"];
		$ret["FAMILYNAME"] = $word["surface"];
		$ret["FAMILYNAMEPHN"] = $word["reading"];
		$ret["GIVENNAME"] = "";
		$ret["GIVENNAMEPHN"] = "";
		break;
	default: // Devided partially
		foreach($array["ma_result"]["word_list"]["word"] as $word){
			if($word["pos"] == "名詞"){
				if(isset($ret["FAMILYNAME"]) == false){
					$ret["FAMILYNAME"] = $word["surface"];
					$ret["FAMILYNAMEPHN"] = $word["reading"];
				} else {
					$ret["GIVENNAME"] .= $word["surface"];
					$ret["GIVENNAMEPHN"] .= $word["reading"];
				}
			}
		}
		break;
	}
	return $ret;
}

// Devide Address into Prefecture, City, Town and Buildings
function parse_address($address){
	global $pref_ar;
	global $dir_req_cache;
	// Get pref level
	foreach($pref_ar as $ac => $pref){ // AreaCode => Name
		if(strncmp($pref, $address, strlen($pref)) == 0){
			$ret["PREF"] = $pref;
			break;
		}
	}
	$ac = sprintf("%02d", $ac);
	if(array_key_exists($ac, $dir_req_cache) == false){
		$buf = file(sprintf(DIR_REQ_URL, $ac));
		$xml = simplexml_load_string(implode("", $buf));
		$json = json_encode($xml);
		$dir_req_cache[$ac] = json_decode($json, TRUE);
	}
	// Get City level
	foreach($dir_req_cache[$ac]["Feature"]["Property"]["AddressDirectory"] as $city){
		$tmp = $ret["PREF"] . $city["Name"];
		$tmp_len = strlen($tmp);
		if(strncmp($tmp, $address, $tmp_len) == 0){
			$ret["ADDRESS"] = $city["Name"];
			$ac = sprintf("%05d", $city["AreaCode"]);
		}
	}
	if(array_key_exists($ac, $dir_req_cache) == false){
		$buf = file(sprintf(DIR_REQ_URL, $ac));
		$xml = simplexml_load_string(implode("", $buf));
		$json = json_encode($xml);
		$dir_req_cache[$ac] = json_decode($json, TRUE);
	}
	// Get Town level
	foreach($dir_req_cache[$ac]["Feature"]["Property"]["AddressDirectory"] as $town){
		$tmp = $ret["PREF"] . $ret["ADDRESS"] . $town["Name"];
		$tmp_len = strlen($tmp);
		if(strncmp($tmp, $address, $tmp_len) == 0){
			$ret["ADDRESS"] .= $town["Name"];
		}
	}
	if(preg_match("/" . $ret["PREF"] . $ret["ADDRESS"] . "([0-9\-]+)(.+)/", $address, $match)){
		$ret["ADDRESS"] .= $match[1];
		$ret["BUILDING"] = trim($match[2]);
		// FIXME
		// I don't know reason why preg_match doesn't pick up last numerics, e.g. "4-1-14" is sometimes divided into "4-1-1" and "4"
		if(preg_match("/^[0-9]+$/", $ret["BUILDING"])){
			$ret["ADDRESS"] .= $ret["BUILDING"];
			$ret["BUILDING"] = "";
		}
	} else {
		$pref_len = strlen($ret["PREF"]);
		$ret["ADDRESS"] = substr($address, $pref_len, strlen($address) - $pref_len + 1);;
		$ret["BUILDING"] = "";
	}
	return $ret;
}

// check path to input
if(!isset($_SERVER["argv"][1])){
	echo("No input file is specified!\n");
	usage();
}

// check path to output
if(!isset($_SERVER["argv"][2])){
	echo("No output directory is specified!\n");
	usage();
}

// check existance of file
$input = realpath($_SERVER["argv"][1]);
if(!file_exists($input)){
	echo("Can't locate specified file!\n");
	usage();
}

// check existance of directory
$dir = realpath($_SERVER["argv"][2]);
if(!file_exists($dir)){
	echo("Can't locate specified directory!\n");
	usage();
}

// check whether $dir is a directory
if(!is_dir($dir)){
	echo("Specified path seems a simple file, NOT a directory!\n");
	usage();
}

// main routine
// read CSV
if(($fp_in = fopen($input, "r")) !== FALSE){
	$line_cnt = 0;
	while(($data = fgetcsv($fp_in, 2048, ",")) !== FALSE){
		$line_cnt += 1;
		if($line_cnt > SKIP_LINE_NO){
			$fields = array();
			if(strlen($data[$csv_field_order["FULLNAME"]]) == 0){
				echo("Line #". $line_cnt . " was skipped due to lack of fullname field.\n");
				continue;
			}
			$ret = get_morphological_analysis($data[$csv_field_order["FULLNAME"]], $fields);
			$fields["FAMILYNAME"] = $ret["FAMILYNAME"];
			$fields["FAMILYNAMEPHN"] = $ret["FAMILYNAMEPHN"];
			$fields["GIVENNAME"] = $ret["GIVENNAME"];
			$fields["GIVENNAMEPHN"] = $ret["GIVENNAMEPHN"];
			$fields["FULLNAME"] = $ret["GIVENNAME"] . " " . $ret["FAMILYNAME"];
			$zip = str_replace("-", "", $data[$csv_field_order["ZIP"]]);
			if(strlen($zip) == 7){ // Avoid oversea ZIP code
				$zip = substr($zip, 0, 3) . "-" . substr($zip, 3, 4);
				$ret = parse_address($data[$csv_field_order["ADDRESS"]]);
				$fields["PREF"] = $ret["PREF"];
				$fields["ADDRESS"] = $ret["ADDRESS"];
				$fields["BUILDING"] = $ret["BUILDING"];
			} else {
				$fields["PREF"] = "";
				$fields["ADDRESS"] = $data[$csv_field_order["ADDRESS"]];
				$fields["BUILDING"] = "";
			}
			$fields["ZIP"] = $zip;
			$fields["ORG"] = $data[$csv_field_order["ORG"]];
			$fields["DEPT"] = $data[$csv_field_order["DEPT"]];
			$fields["TITLE"] = $data[$csv_field_order["TITLE"]];
			$fields["EMAIL"] = $data[$csv_field_order["EMAIL"]];
			$fields["TEL"] = str_replace("-", "", $data[$csv_field_order["TEL"]]);
			$fields["URL"] = $data[$csv_field_order["URL"]];

			$vcf_cont = VCF_TMPLT;
			foreach($fields as $key => $value){
				$vcf_cont = str_replace('{'.$key.'}', $value, $vcf_cont);
			}
			$tmp_path = sprintf("%s/%s %s", $dir, $fields["FAMILYNAME"], $fields["GIVENNAME"]);
			$num_of_files = count(glob($tmp_path . "*"));
			if($num_of_files != 0){
				$output = $tmp_path . sprintf("_%d.vcf", $num_of_files + 1);
			} else {
				$output = $tmp_path . ".vcf";
			}
			if(($fp_out = fopen($output, "w")) !== FALSE){
				fwrite($fp_out, $vcf_cont);
				fclose($fp_out);
				echo("Line #". $line_cnt . " was written.\n");
			}
		}
	}
	fclose($fp_in);
}

?>
