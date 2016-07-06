<?php
/** 
 * Excel Map is a tool to parse and look up values from an excel sheet
 * @package excel2JSON
 * @license http://www.apache.org/licenses/LICENSE-2.0
 * @author Karl Holz
 * @version 1.0
 */
class excel2JSON  { 

	/**
	 * excel file path
	 * @var string $file
	 */
	private $file=FALSE;

	/**
	 * excel file with zip:// prefix
	 * @var string $excel
	 */
	private $excel=FALSE;

	/**
	 * excel workbooks reference
	 * @var string $workbooks
	 */
	private $workbooks="xl/workbook.xml";

	/**
	 * excel XML shared strings reference file
	 * @var string $strings
	 */
	private $strings="xl/sharedStrings.xml";

	/**
	 * excel shared Strings data
	 * @var string $excel_strings
	 */
	private $excel_strings=FALSE;
	
	/**
	 * loaded excel file list
	 * @var array $load
	 */
	private $load=array();
	
	/** 
	 * selected Excel workbook
	 * 
	 * @var string $loaded_workbook
	 */
	private $loaded_workbook=FALSE;

	
	/**
	 * Loaded Workbook data
	 * 
	 * @var array $loaded_workbook_data
	 */
	private $loaded_workbook_data=array();
	
	function get_loaded_workbook_data() { return $this->loaded_workbook_data ;}

	/**
	 * Loaded Workbook data
	 *
	 * @var array $new_loaded_workbook_data
	 */
	private $new_loaded_workbook_data=array();

	function get_new_loaded_workbook_data() { return $this->new_loaded_workbook_data ;}
	
	/**
	 * loaded Workbook data
	 * 
	 * @var array $loaded_workbook_cells
	 */
	private $loaded_workbook_cells=array();
	
	function get_loaded_workbook_cells() { return $this->loaded_workbook_cells ;}

	/**
	 * loaded Workbook data
	 *
	 * @var array $loaded_workbook_rows
	 */
	private $loaded_workbook_rows=array();
	
	function get_loaded_workbook_rows () { return $this->loaded_workbook_rows ;}

	/**
	 * loaded Workbook data
	 *
	 * @var array $loaded_workbook_cols
	 */
	private $loaded_workbook_cols=array();

	function get_loaded_workbook_cols () { return $this->loaded_workbook_cols ;}
	
	/**
	 * Excel cell list
	 * @var array $filter
	 */
	private $filter=array();


	/**
	 * load the excel file if it's past as an argument
	 *
	 * @param string $file
	 */
	function __construct($file='') {
		if ($file != '')
			$this->set_excel($file);
		return TRUE;
	}

	private $ini=FALSE;

	function load_config() {
		$ini=basename($this->excel, '.xlsx') . '.ini';
		if (!is_file($ini)) {
			foreach ($this->get_sheets() as $p){
				$this->load_sheet($p);
				$this->load_sheet_data();
				$excel=$this->get_loaded_workbook_cells();
				$ini.=<<<i
[$p]
i
				;
				foreach ($excel as $k =>$v)	$ini.="$k=\"$k\"\n";
			}
			/**
			 * write configuration file as test.ini
			 */
			file_put_contents(basename($excel, '.xlsx') . '.ini', $ini);
		}
		$this->ini=parse_ini_file($ini, TRUE);
		return TRUE;
	}

	function get_config() {
		if (array_key_exists($this->sheet, $this->ini)) return $this->ini[$this->sheet];
		return FALSE;
	}

	function get_sheets() {
		$path=array();
		/**
		 * get a list of excel sheets and add them to $path
		 */
		if (!$this->file) return $path;
		foreach ($this->excel_sheets['title'] as $k => $v) $path[]=$v;
		return $path;
	}

	/**
	 * Set Excel doc
	 * Sets the excel file to use as a DB
	 * file should have an xlsx or xlsm
	 * @param string $file
	 * @param string $t it's default is true, set to false if your sheet fails to load
	 * @return boolean
	 */
	function set_excel($file, $t='true') { 
		if (is_file($file)) { 
			$this->excel=realpath($file);
			$this->file='zip://'.realpath($file);
			$this->load=$this->zip_resorces($this->excel);
			// load work book list
			$sheets=trim(file_get_contents($this->file.'#'.$this->workbooks));
			$wb=$this->xsl_out($this->GetWorkbooksList_xslt($t), $sheets);
			if ($wb != '') $this->excel_sheets=parse_ini_string($wb);
			$this->load_config();
			return TRUE;
		} 
		return FALSE; 
		
	}

	private $sheet = FALSE;

	/**
	 * Load Excel sheet by name
	 * @param string $sheet
	 * @return boolean
	 */
	function load_sheet($sheet) {
		$this->sheet=$sheet;
		if (in_array($sheet, $this->excel_sheets['title'])) {
			$key=array_search($sheet, $this->excel_sheets['title']);
			$this->loaded_workbook=ltrim($this->excel_sheets['xml'][$key], '/');
			return TRUE;
		}
		return FALSE;
	}



	/**
	 * Load Excel sheet data
	 * @return boolean
	 */
	function load_sheet_data(){
		if (!$this->loaded_workbook) return FALSE;

		$ini=$this->xsl_out($this->cellList_xslt(), $this->get_workbook());
		if ($ini != '') {
			$this->loaded_workbook_data=parse_ini_string($ini, TRUE);
			if(array_key_exists('sheet', $this->loaded_workbook_data)) {
				if(array_key_exists('dimension', $this->loaded_workbook_data['sheet'])) {
					$this->col_start = $this->get_col($this->loaded_workbook_data['sheet']['dimension'][0]);
					$this->col_end   = $this->get_col($this->loaded_workbook_data['sheet']['dimension'][1]);
					$this->row_start = $this->get_row($this->loaded_workbook_data['sheet']['dimension'][0]);
					$this->row_end   = $this->get_row($this->loaded_workbook_data['sheet']['dimension'][1]);
				}
			}
			array_shift($this->loaded_workbook_data);
			//setup excel cell filter



			foreach($this->loaded_workbook_data as $k => $v) {
				if (array_key_exists($k, $this->ini[$this->sheet])) {
					$label=$this->ini[$this->sheet][$k];
				} else {
					$label=FALSE;
				}

				if (array_key_exists('val', $v)) {
					if (array_key_exists('t', $v) && $v['t'] == 's') {
						$val=$v['val']+1; //add 1 so search matches value in excel_strings
						$s=$this->xsl_out($this->StringLookup_xslt( $val), $this->get_excel_strings());

						if ($label) $this->new_loaded_workbook_data[$label]['val']=$s;

						$this->loaded_workbook_data[$k]['val']=$s;

						$this->loaded_workbook_cells[$k]=$s;

						$this->loaded_workbook_rows[$this->get_row($k)][$this->get_col($k)]=$s;

						$this->loaded_workbook_cols[$this->get_col($k)][$this->get_row($k)]=$s;

						// drop the 't' key off of the array, it's only needed for text look up
						array_pop($this->loaded_workbook_data[$k]);
					}  else {

						if ($label) $this->loaded_workbook_cells[$label]=$v['val'];

						$this->loaded_workbook_rows[$this->get_row($k)][$this->get_col($k)]=$v['val'];

						$this->loaded_workbook_cols[$this->get_col($k)][$this->get_row($k)]=$v['val'];
					}



					// new workbook data with custom label
					if ($label){
						$this->new_loaded_workbook_data[$label]['col']=$this->get_col($k);
						$this->new_loaded_workbook_data[$label]['row']=$this->get_row($k);
						$this->new_loaded_workbook_data[$label]['cell']=$k;

					}


					//updated workbook data
					$this->loaded_workbook_data[$k]['col']=$this->get_col($k);
					$this->loaded_workbook_data[$k]['row']=$this->get_row($k);
					$this->loaded_workbook_data[$k]['cell']=$k;
					if ($label) $this->loaded_workbook_data[$k]['label']=$label;

				}

			}

			foreach ($this->ini[$this->sheet] as $k => $v) {
				$this->add_cell_list($k);
			}

			//	echo '<pre>';print_r($this->ini[$this->sheet]); echo '</pre>'; exit();

			return TRUE;
		}
		return FALSE;
	}

	private $col_start = 'A';

	private $col_end = 'ZZ';

	private $row_start = 1;

	private $row_end = 512;


	/**
	 * get the loaded workbook page
	 * @return string
	 */
	function get_workbook() {
		return trim(file_get_contents($this->file.'#'.$this->loaded_workbook));
	}

	/**
	 * get Excel strings from the selected file
	 *
	 * @return string
	 */
	function get_excel_strings() {
		return trim(file_get_contents($this->file.'#'.$this->strings));
	}

	/**
	 * returns the colum used in excel
	 *
	 * example A2 will return A
	 *
	 * @param $cell
	 * @return string
	 */
	function get_col($cell) {
		return preg_replace('/\d*/', '', $cell);
	}

	/**
	 * returns the row used in excel
	 *
	 * example B6 will return 6
	 *
	 * @param $cell
	 * @return int
	 */
	function get_row($cell) {
		return preg_replace('/[A-Za-z]*/', '', $cell);
	}

	/**
	 * Cell List
	 * add an excel cell reference to filter list
	 * @param string $cell
	 * @return bool
	 */
	function add_cell_list($cell) {
		if (array_key_exists($cell, $this->loaded_workbook_cells)) $this->filter[]=$cell;
		return TRUE;
	}
	
	/**
	 * Print Rows
	 * print sheet cell rows as JSON Document, it will basicly be { "excel cell":"excel data" }
	 * @return bool
	 */
	function print_sheet_rows() {
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->loaded_workbook_rows);
		return TRUE;
	}
	
	/**
	 * Print Sheet
	 * print sheet data as JSON Document, it will basicly be { "excel cell" : { "row": "row number", "val": "excel data" } }
	 * @return bool
	 */
	function print_sheet_data() { 
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->loaded_workbook_data);
		return TRUE;
	}
	
	/**
	 * Filter Data
	 * Filter workbook data 
	 * @return array filter data
	 */
	function filter_data() {
		$filter=array();
		foreach ($this->filter as $f) $filter[]=$this->loaded_workbook_data[$f];
		return $filter;
	}
	
	
	/**
	 * Print filtered sheet
	 * Print filtered sheet cells data as JSON Document, it will basicly be { "excel cell" : { "row": "row number", "val": "excel data" } }
	 *
	 */
	function print_filter_data() {
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->filter_data());
	}
	
	
	/**
	 * Print Cells
	 * print sheet cells as JSON Document, it will basicly be { "excel cell":"excel data" }
	 * 
	 */
	function print_sheet_cells() {
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->loaded_workbook_cells);
	}
	
	/**
	 * Filter cells 
	 * Returns filter data as an array
	 * @return array
	 */
	function filter_cells() {
		$filter=array();
		foreach ($this->filter as $f)
			if (array_key_exists($f, $this->ini[$this->sheet]))
				$filter[$this->ini[$this->sheet][$f]]=$this->loaded_workbook_cells[$f];
		return $filter;
	}
	
	/**
	 * Print filtered cells
	 * Print filtered sheet cells as JSON Document, it will basicly be { "excel cell":"excel data" }
	 *
	 * @return bool
	 */
	function print_filter_cells() {
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->filter_cells());
		return TRUE;
	}
	
	/**
	 * List Workbooks
	 * print workbook sheet names as JSON Document
	 *
	 * it will basicly be [ "sheet name" ]
	 *
	 * @return bool
	 */
	function print_sheets() {
		$this->json_out($this->excel_sheets['title']);
		return TRUE;
	}
	
	/**
	 * Open Zip
	 * Open files such as excel xlsx or xlsm
	 * @param $zip excel doc
	 * @return array
	 */
	private function zip_resorces($zip) {
		if (! class_exists ( 'ZipArchive' )) die('ZipArchive class not found');
		$za = new ZipArchive();
		$za->open( $zip );
		$list = array();
		for($i = 0; $i < $za->numFiles; $i ++) {
			$z = $za->statIndex( $i );
			$list[] = $z['name'];
		}
		$za->close();
		return $list;
	}
	
	/**
	 * Print JSON
	 * 
	 * Print out json document to browser/client
	 * @param array $data array to convert to JSON 
	 * @return string
	 */
	private function json_out($data) {
		if (! is_array ( $data ) || (count ( $data ) == 0)) $data=array("error" => "No Excel data to show");
		header ( 'Cache-Control: no-cache, must-revalidate' );
		header ( 'Expires: Mon, 26 Jul 1997 05:00:00 GMT' );
		header ( 'Content-type: application/json' );
		echo json_encode ( $data );
		exit ();
	}
	
	/**
	 * XSLT out
	 * 
	 * Apply xslt template to xml data
	 *
	 * @param string $xsltmpl         	XSLT stylesheet to be applied to XML
	 * @param string $xml_load      	XML data
	 * @return string 
	 */
	private function xsl_out($xsltmpl, $xml_load) {
		if (! class_exists ( 'DOMDocument' ))  die('DOMDocument class not found');
		if (! class_exists ( 'XSLTProcessor' )) die('XSLTProcessor class not found');
		// loads XML data string
		$xml = new DOMDocument ();
		if (! $xml->loadXML ( $xml_load )) die('XML data failed to load'); 
		// loads XSL template string
		$xsl = new DOMDocument ();
		if (! $xsl->loadXML ( $xsltmpl ))  die('XSLT failed to load'); 
		$xslproc = new XSLTProcessor();
		$xslproc->importStylesheet( $xsl );
		return $xslproc->transformToXml( $xml );
	}

/* XSLT TEMPLATES THAT WILL BE USED TO PARSE XML IN THE EXCEL ZIP FILE */
	
	/**
	 * Select all cells
	 * 
	 * XSLT style for selecting all cells on an excel work sheet
	 * XML: <excel>/xl/worksheets/sheet*.xml 
	 * the xml file needed is sourced from the result from the output from the method bellow
	 * @see excel2JSON::GetWorkbooksList_xslt()
	 * @return string
	 */
	private function cellList_xslt() {
		return <<<X
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
 xmlns:w="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
 <xsl:output method="text" media-type="text/plain"/>
 <xsl:template match="/">[sheet]
dimension[]="<xsl:value-of select="substring-before(//dimension/@ref|//w:dimension/@ref,':')"/>"
dimension[]="<xsl:value-of select="substring-after(//dimension/@ref|//w:dimension/@ref,':')"/>"
<xsl:apply-templates select="//c|//w:c" /></xsl:template>
 <xsl:template match="//c|//w:c">
  <xsl:if test="v|w:v">
[<xsl:value-of select="@r"/>]
; row="<xsl:value-of select="../@r"/>"
val="<xsl:value-of select="v|w:v"/>"
<xsl:if test="@t">t="<xsl:value-of select="@t"/>"</xsl:if>
  </xsl:if>
 </xsl:template>
</xsl:stylesheet>
X
	;
	}
	
	/**
	 * XSLT style string lookup
	 * 
	 * Search for string values by row value
	 * XSLT param is called row, should be numeric, but can be overriden by passing a numeric paramater
	 * XML: <excel>/xl/sharedStrings.xml
	 * @param string $row row can be the default, or the number to look up.  this is a work around
	 * @return string
	 */
	private function StringLookup_xslt($row='$row') {
		return <<<x
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 
	xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	version="1.0">
  <xsl:output method="text" media-type="text/plain"/>
  <!-- Row should be numeric, starts at 1, not 0; may need to add 1 to row first -->
  <xsl:param name="row" />
  <xsl:template match="/">
    <xsl:value-of select="/sst/si[$row]/t|/s:sst/s:si[$row]/s:t"/></xsl:template>
</xsl:stylesheet>
x
;
	}
	
	/**
	 * Get Workbooks
	 * 
	 * XSLT style for grabing a list of the excel worksheets
	 * XML: <excel>/xl/workbook.xml
	 * @param string $t true by default, otherwise false for all
	 * @return string
	 */
	private function  GetWorkbooksList_xslt($t='true') {
		$page='$page';
		$r='$r';
		return <<<x
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
  xmlns:o="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  version="1.0">
  <xsl:output method="text" media-type="text/plain"/>
  <xsl:param name="page" select="'xl/worksheets/sheet'" />
  <xsl:param name="r" select="'$t'"/>
  <xsl:template match="/"><xsl:apply-templates select="o:workbook|workbook" />
 </xsl:template>
  <xsl:template  match="o:workbook|workbook">[worksheets]<xsl:apply-templates select="o:sheets|sheets" /></xsl:template>
  <xsl:template match="o:sheets|sheets">
    <xsl:for-each select="o:sheet|sheet">
title[]="<xsl:value-of select="@name"/>"
<xsl:choose>
 <xsl:when test="$r='true'">
xml[]="<xsl:value-of select="$page"/><xsl:value-of select="substring-after(@r:id, 'rId')"/>.xml"   
  </xsl:when>
  <xsl:otherwise>
xml[]="<xsl:value-of select="$page"/><xsl:value-of select="@sheetId"/>.xml"
  </xsl:otherwise>
</xsl:choose>  
  </xsl:for-each>
 </xsl:template>
</xsl:stylesheet>
x
;
	}
	
}

/* Demo file and sheets */
$file='test.xlsx';
$sheet='Filelist';
