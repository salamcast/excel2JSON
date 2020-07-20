<?php
require_once 'excelMap.class.php';
$e=new excel2JSON();
//$e->set_excel($file);
$e->set_excel($file, 'false');
$e->print_sheets();