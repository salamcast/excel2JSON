<?php
require_once 'excelMap.class.php';
$e=new excel2JSON();
/**
 * you will need to set the $file variable for this to work
 */
$e->set_excel($file, 'false');
$path=array();
/**
 * get a list of excel sheets and add them to $path
 */
foreach ($e->excel_sheets['title'] as $k => $v) $path[]=$v;
//
$ini='';
/**
 * make the ini configuration for each sheet, 
 * only file and sheet, plus the cells with data are set as keys
 */
foreach ($path as $p){
	unset($e);
	$e=new excel2JSON();
	$e->set_excel($file, 'false');
	$e->load_sheet($p);
	$e->load_sheet_data();
	$excel=$e->get_loaded_workbook_cells();
//	print_r($excel);
	$ini.=<<<i
<<<<<<< HEAD
[/$p]
file="$file"
sheet="$p"

=======
[$p]
>>>>>>> 8f4f2cc92a9e44e1db86f9518cc638864b8b70d0
i
;
	foreach ($excel as $k =>$v)	$ini.="$k=\"$k\"\n";
}
/**
<<<<<<< HEAD
 * write configuration file as config.ini
 */
file_put_contents('config.ini', $ini);
=======
 * write configuration file as test.ini
 */
file_put_contents(basename($file, '.xlsx') . '.ini', $ini);
>>>>>>> 8f4f2cc92a9e44e1db86f9518cc638864b8b70d0
/**
 * show newly made file
 */
if (basename($_SERVER['SCRIPT_NAME']) == 'makeconfig.php')echo "<pre>".$ini."</pre>";
