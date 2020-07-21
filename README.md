# Excel2JSON is an Excel XML parser to JSON document 

This class will allow you to use your excel workbook as a backend for your Javascript/jQuery and other web applications.
It's really simple to use, you can instantiate the class with the excel file you want to read, the file must 
be ooxml, normally with an xlsx or xlsm file extention MS Excel 2007
 
### Won't work with excel (xls) docs exported with:
 - Apache Open office (excel 2003 XML save as clashes)
 - Numbers 
 - Libre Office (excel 2003 XML save as clashes)
 
I made test.xlsx with Oxygen XML, and edited it with LibreOffice. it is the test excel file provided with this class
   
*** To see this class in action, view webpage.php and view it's code.

Note: if you get an error with the JSON output, set_excel with a second argument as seen bellow
```
    $e->set_excel($file, 'false');
```
other wise use 
```
    $e->set_excel($file);
```
# configuration support

The makeconfig.php will generate an ini file so you can filter out cells and change their label, like A1 to Name and A2 the something like surname.
Take a look at the example bellow or at the config.ini for more details.
```
[/Filelist]
file="test.xlsx"
sheet="Filelist"
A1="title"
A2="cell_label"
B2="cell"
[/ty]
file="test.xlsx"
sheet="ty"
A1="Apply"
K10="Keep"
E16="Empty"
L56="List"
```
# Examples

### Sheets test - print_sheets.php => on by default if worksheet not found or set.  

Prints out a JSON document of a list of excel workbook sheets

```
    require_once 'excelMap.class.php';
    $e=new excel2JSON();
    //$e->set_excel($file);
    $e->set_excel($file, 'false');
    $e->print_sheets();
```

### Data test - print_sheet_data.php

Prints out a JSON document of excel worksheet cell data in detail

```
    require_once 'excelMap.class.php';
    $e=new excel2JSON();
    //$e->set_excel($file);
    $e->set_excel($file, 'false');
    $e->load_sheet($sheet);
    $e->load_sheet_data();
    $e->print_sheet_data();
```

### Cell test - print_sheet_cells.php

Prints out a JSON document of excel worksheet cell and their values 

```
    require_once 'excelMap.class.php';
    $e=new excel2JSON();
    //$e->set_excel($file);
    $e->set_excel($file, 'false');
    $e->load_sheet($sheet);
    $e->load_sheet_data();
    $e->print_sheet_cells();
```

### Rows test - print_sheet_rows.php

Prints out a JSON document of excel worksheet cell and values grouped into rows 

```
    require_once 'excelMap.class.php';
    $e=new excel2JSON();
    //$e->set_excel($file);
    $e->set_excel($file, 'false');
    $e->load_sheet($sheet);
    $e->load_sheet_data();
    $e->print_sheet_rows();
```


### Data filter test - print_filter_data.php

Prints out a JSON document of excel worksheet cell data in detail, only cells added to cell list will be displayed

```
    require_once 'excelMap.class.php';
    $e=new excel2JSON();
    //$e->set_excel($file);
    $e->set_excel($file, 'false');
    $e->load_sheet($sheet);
    $e->load_sheet_data();
    $e->add_cell_list('A1');
    $e->add_cell_list('A2');
    $e->add_cell_list('B2');
    $e->add_cell_list('A4');
    $e->add_cell_list('B4');
    $e->print_filter_data();
```


### Cell filter test - print_filter_cells.php

Prints out a JSON document of excel worksheet cell and their values, only cells added to cell list will be displayed 

```
    require_once 'excelMap.class.php';
    $e=new excel2JSON();
    //$e->set_excel($file);
    $e->set_excel($file, 'false');
    $e->load_sheet($sheet);
    $e->load_sheet_data();
    $e->add_cell_list('A1');
    $e->add_cell_list('A2');
    $e->add_cell_list('B2');
    $e->add_cell_list('A4');
    $e->add_cell_list('B4');
    $e->print_filter_cells();
```