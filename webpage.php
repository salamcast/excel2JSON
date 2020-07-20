<?php
require_once 'excelMap.class.php';
/**
 * make a configuration file
 */
if (! is_file('config.ini')) require_once 'makeconfig.php';
$ini=parse_ini_file('config.ini', TRUE);
$dir=dirname($_SERVER['SCRIPT_NAME']);
$error=array();


/**
 * debug_panel html markup
 * @param unknown_type $title
 * @param unknown_type $json
 * @return string
 */
function debug_panel($title, $json){
	$dump=print_r(json_decode($json,true),true);

	return <<<h
			<div class="col-md-4 panel panel-default">
 				<div class="panel-heading">$title</div>
 					<div class="panel-body">
 						$json<hr/><pre>$dump</pre>
 					</div>
				</div>	
			</div>
h
	;
}
/**
 * debug data table html
 * @param unknown_type $title
 * @param unknown_type $data
 * @param unknown_type $check
 * @return string
 */
function debug_data_table($title,$data,$check){
	$tbl=<<<t
			<div class="col-md-4 panel panel-default">
				<div class="panel-heading">$title</div>
					<div class="panel-body">
						<table class="table">
							<tr><th>Excel</th><th>JSON</th><th>Data</th></tr>
t
	;
	foreach ($data as $k => $v) {
		$x="*";
		if (array_key_exists($k, $check)) { $x=$check[$k];
			$tbl.= <<<t
							<tr><td>$k</td><td>$v</td><td>$x</td></tr>
t
			;
		}		
	}
	$tbl.=<<<t
					 	</table>
					</div>
				</div>
			</div>
t
	;
	return $tbl;
}

/**
 * make nav bar of all configured excel sheets mapped to a uri
 * @param unknown_type $b
 * @param unknown_type $ini
 * @return string
 */
function make_nav($b, $ini){
	$n=<<<h
			<nav class="navbar navbar-default" role="navigation">
  				<div class="navbar-header"<a class="navbar-brand" href="#">$b</a></div>
  					<div class="collapse navbar-collapse" id="excel_nav">
						<ul class="nav navbar-nav">
h
	;
	//print out link list
	$srv=$_SERVER['SCRIPT_NAME'];
	foreach ($ini as $k => $v)
		$n.=<<<h
							<li><a href="$srv$k" >$k</a></li>
h
		;
	
	$n.=<<<h
						</ul>
				</div>
			</nav>
h
	;
	return $n;
}
// ==============================================================

//if not path_info, then stop, nothing more can be done

if ( array_key_exists('PATH_INFO', $_SERVER)) {
	$path=$_SERVER['PATH_INFO'];
	
	
	
	// print error page if path is not configured
	if (! array_key_exists($path, $ini))
		$error[]='<div class="alert alert-danger">ERROR: 404 Page not found: '.$path.'</div>';
	//check for file,send error
	if (! array_key_exists('file', $ini[$path]) || ! is_file($ini[$path]['file']))
		$error[]='<div class="alert alert-danger">ERROR: file is not configured or not found for: '.$path.'</div>';
	
	$file=$ini[$path]['file'];
	
	if (! array_key_exists('sheet', $ini[$path]))
		$error[]='<div class="alert alert-danger">ERROR: sheet is not configured for: '.$path.'</div>';
	
	$sheet=$ini[$path]['sheet'];
} else {
	$path=false;
	$sheet=false;
	$file=false;
}
?><!DOCTYPE html>
<html>
	<head>
		<title>Excel2JSON Debug viewer webpage</title>
	    <meta name="viewport" content="width=device-width, initial-scale=1.0">
		<link rel="stylesheet" href="<?php echo $dir; ?>/dist/css/bootstrap.min.css">
		<link rel="stylesheet" href="<?php echo $dir; ?>/dist/css/bootstrap-theme.min.css">
		<script src="<?php echo $dir; ?>/dist/js/jquery.min.js"></script>
		<script src="<?php echo $dir; ?>/dist/js/bootstrap.min.js"></script>
	</head>
	<body>
<?php 
echo make_nav('Excel2JSON', $ini);

if (count($error) > 0){
	foreach ($error as $er) echo $er;
} else {
//load excel class to process
	$e=new excel2JSON();
	$e->set_excel($file, 'false');
	$e->load_sheet($sheet);
	$e->load_sheet_data();

//setup excel cell filter
	$filter=array();
	foreach ($ini[$path] as $k => $v) {
		if ($k != 'sheet' || $k != 'file') {
			$filter['s'][]='/\"'.$k.'\"/';
			$filter['r'][]='"'.$v.'"';
			$filter['data'][$k]=$v;
			$e->add_cell_list($k);
		}
	}

	//excel sheet data
	$excel=$e->filter_cells();

	$cells_title='Filtered Cells | excel2JSON::filter_cells()';
	$cells=preg_replace($filter['s'], $filter['r'], json_encode($excel), 1);

	$data_title='Filtered Data | excel2JSON::filter_data()';
	$data=preg_replace($filter['s'], $filter['r'], json_encode($e->filter_data()));
	
	$cells2_title='All Cells | excel2JSON::get_loaded_workbook_cells()';
	$cells2=preg_replace($filter['s'], $filter['r'], json_encode($e->get_loaded_workbook_cells()), 1);
	
	$data2_title='All Data | excel2JSON::get_loaded_workbook_data()';
	$data2=preg_replace($filter['s'], $filter['r'], json_encode($e->get_loaded_workbook_data()));
	
	$rows_title='All Rows | excel2JSON::get_loaded_workbook_rows()';
	$rows=preg_replace($filter['s'], $filter['r'], json_encode($e->get_loaded_workbook_rows()), 1);
	
	//////////
	?>
			<div class="row">
	<?php 
	//JSON output from Excel2JSON class
	echo debug_data_table('Filtered Cells View', $filter['data'], $excel);
	echo debug_data_table('Workbook Data', $filter['data'], $e->get_loaded_workbook_cells());
	?>
			</div>
			<div class="row">
	<?php
	echo debug_panel($cells_title, $cells);
	echo debug_panel($data_title, $data);
	?>
			</div>
			<div class="row">
	<?php
	echo debug_panel($cells2_title, $cells2);
	echo debug_panel($data2_title, $data2);
	echo debug_panel($rows_title,$rows);
	?>
			</div>
	<?php
}
?>
	</body>
</html>
