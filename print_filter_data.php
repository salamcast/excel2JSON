<?php
require_once 'excelMap.class.php';
$e=new excel2JSON();
$e->set_excel($file, 'false');
$e->load_sheet($sheet);
$e->load_sheet_data();
//
$e->add_cell_list('A1');

$e->add_cell_list('A2');
$e->add_cell_list('B2');

$e->add_cell_list('A4');
$e->add_cell_list('B4');

$e->print_filter_data();