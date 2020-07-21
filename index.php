<?php
require_once 'excelMap.class.php';
/**
 * make a configuration file
 */
$conf='test.ini';
$dir=dirname($_SERVER['SCRIPT_NAME']);
$error=array();


/**
 * debug_panel html markup
 * @param unknown_type $title
 * @param unknown_type $json
 * @return string
 */
	/*

	 */
function debug_panel($title, $json, $box=0){
	$dump=print_r(json_decode($json,true),true);

	return <<<h
		<div class="panel-group">
			<div class="panel panel-default">
				<div class="panel-heading">
					<h4 class="panel-title">
						<a data-toggle="collapse" href="#collapse$box">$title</a>
					</h4>
				</div>
				<div id="collapse$box" class="panel-collapse collapse">
					<div class="panel-body">
						<h5>JSON Output</h5>
						<code>$json</code>
					</div>
					<div class="panel-footer">
						<h5>print_r dump</h5>
						<pre class="pre-scrollable" >$dump</pre>
					</div>
				</div>
			</div>
		</div>
h
	;
}


/**
 * make nav bar of all configured excel sheets mapped to a uri
 * @param unknown_type $b
 * @param unknown_type $ini
 * @return string
 */
function make_nav($b, $ini){
	$srv=$_SERVER['SCRIPT_NAME'];
	$n=<<<h
			<nav class="navbar navbar-inverse">
  				<div class="container-fluid">
    				<div class="navbar-header"><a class="navbar-brand" href="$srv">$b</a></div>
					<ul class="nav navbar-nav">
h
	;
	//print out link list

	foreach ($ini as $k )
		$n.=<<<h
						<li><a href="$srv/$k" >$k</a></li>
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


$excel=basename($conf, '.ini') . ".xlsx";
$e=new excel2JSON();
if (! $e->set_excel($excel, 'false'))
	$error[]='<div class="alert alert-danger">ERROR: file is not configured or not found for: '.$excel.'</div>';

if ( array_key_exists('PATH_INFO', $_SERVER)) {
	$path=trim($_SERVER['PATH_INFO'], '/');


	if (count($error) == 0){

		if (!$e->load_sheet($path))
			$error[]='<div class="alert alert-danger">ERROR: 404 Page not found: '.$path.'</div>';
		//check for file,send error

		$e->load_sheet_data();

	}


} else {
	$path=false;
	$sheet=false;
}
?><!DOCTYPE html>
<html>
	<head>
		<title>Excel2JSON Debug viewer webpage</title>
	    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
		<link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.2/jquery.min.js"></script>
		<script src="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>
	</head>
	<body>



<?php 
echo make_nav('Excel2JSON', $e->get_sheets());

if (count($error) > 0){
	foreach ($error as $er) echo $er;
} else {
//load excel class to process
	?>


	<?php
	if ($path) {
		$x=0;

		$conf_title='Sheet Conf | excel2JSON::get_config()';
		$conf=json_encode($e->get_config());
		echo debug_panel($conf_title, $conf, $x); $x++;

		$cells_title='Filtered Cells | excel2JSON::filter_cells()';
		$cells=json_encode($e->filter_cells());
		echo debug_panel($cells_title, $cells, $x);$x++;

		$cells2_title='All Cells | excel2JSON::get_loaded_workbook_cells()';
		$cells2=json_encode($e->get_loaded_workbook_cells());
		echo debug_panel($cells2_title, $cells2, $x);$x++;

		$rows_title='All Rows | excel2JSON::get_loaded_workbook_rows()';
		$rows=json_encode($e->get_loaded_workbook_rows());
		echo debug_panel($rows_title,$rows, $x);$x++;

		$cols_title='All Cols | excel2JSON::get_loaded_workbook_cols()';
		$cols=json_encode($e->get_loaded_workbook_cols());
		echo debug_panel($cols_title,$cols, $x);$x++;

		$data_title='Filtered Data | excel2JSON::filter_data()';
		$data=json_encode($e->filter_data());
		echo debug_panel($data_title, $data, $x);$x++;

		$data2_title='All Data | excel2JSON::get_loaded_workbook_data()';
		$data2=json_encode($e->get_loaded_workbook_data());
		echo debug_panel($data2_title, $data2, $x);$x++;

		$new_data2_title='All Data | excel2JSON::get_new_loaded_workbook_data()';
		$new_data2=json_encode($e->get_new_loaded_workbook_data());
		echo debug_panel($new_data2_title, $new_data2, $x);$x++;


	}

}
?>
	</body>
</html>
