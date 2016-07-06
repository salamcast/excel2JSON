<?php
require_once 'excelMap.class.php';
$e=new excel2JSON();
//$e->set_excel($file);
$e->set_excel($file, 'false');
$e->load_sheet($sheet);
$e->load_sheet_data();
$e->print_sheet_data();
