<?php
/** 
 * Excel Map is a tool to parse and look up values from an excel sheet
 * @package excel2JSON
 * @license http://www.apache.org/licenses/LICENSE-2.0
 * @author Abu Khadeejah Karl Holz
 * @version 2.0
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
	private $excel_strings='';
	
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
	 * Excel cell list
	 * @var array $filter
	 */
	private $filter=array();

	/**
	 * Custom Config label
	 * @var array $filter_label
	 */
	private $filter_label=array();
	
	
	/**
	 * load the excel file if it's past as an argument
	 * 
	 * @param string $file
	 * @return boolean
	 */
	function __construct($file='') {
		if ($file != '') return $this->set_excel($file);
		return TRUE;
	}
	
	/**
	 * Set Excel doc
	 * Sets the excel file to use as a DB
	 * file should have an xlsx or xlsm
	 * @param file $file
	 * @param string it's default is true, set to false if your sheet fails to load
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
			return TRUE;
		} 
		return FALSE; 
		
	}
	
	/**
	 * Load Excel sheet by name
	 * @param string $sheet
	 * @return boolean
	 */
	function load_sheet($sheet) {
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
		$data=trim(file_get_contents($this->file.'#'.$this->loaded_workbook));
		$excel_strings=trim(file_get_contents($this->file.'#'.$this->strings));
		$ini=$this->xsl_out($this->cellList_xslt(), $data);
		if ($ini != '') {
			$this->loaded_workbook_data=parse_ini_string($ini, TRUE);
			foreach($this->loaded_workbook_data as $k => $v) {
				if ($k != 'sheet') {
					if (array_key_exists('val', $v)) {
						if (array_key_exists('t', $v) && $v['t'] == 's') {	
							$val=$v['val']+1;
							$s=$this->xsl_out($this->StringLookup_xslt( $val), $excel_strings);
							$this->loaded_workbook_data[$k]['val']=$s;
							// drop the 't' key off of the array, it's only needed for text look up
							array_pop($this->loaded_workbook_data[$k]);
							$this->loaded_workbook_data[$k]['cell']=$k;
							$this->loaded_workbook_cells[$k]=$s;
							$this->loaded_workbook_rows[$v['row']][$k]=$s;

					 	}  else {
					 		$this->loaded_workbook_cells[$k]=$v['val'];
					 		$this->loaded_workbook_rows[$v['row']][$k]=$v['val'];
						}
					}
				}
			}
			return TRUE;
		}
		return FALSE;
	}
	
	public $config_label=array();

	function load_config($ini, $sec) {
		if (! is_file($ini)) return FALSE;
		$ini=parse_ini_file($ini, TRUE);
//setup excel cell filter
		foreach ($ini[$sec] as $k => $v) {
			if ($k != 'sheet' || $k != 'file') {
//				$filter['s'][]='/\"'.$k.'\"/';
//				$filter['r'][]='"'.$v.'"';
				$this->config_label[$k]=$v;
				$this->filter_label[] = $v;
				$this->filter[] = $k;
			}
		}



//		$cellsLoaded=array();
//		foreach ($this->filter as $k => $v) {
//			if (array_key_exists($k, $this->loaded_workbook_cells)) {
//				$key=$filter['data'][$k];
//			} else {
//				$key=$k;
//			}
//			$cellsLoaded[$key]=$v;
//		}
//		$this->loaded_workbook_cells=$cellsLoaded;

		/*$dataLoaded=array();
		foreach($this->loaded_workbook_data as $k => $v) {
			if (array_key_exists($k, $filter['data'])) {
				$key=$filter['data'][$k];
			} else {
				$key=$k;
			}
			$dataLoaded[$key] = $v;
		}
		$this->loaded_workbook_data=$dataLoaded;*/

		return TRUE;
	}

	/**
	 * Cell List
	 * add an excel cell reference to filter list
	 * @param string $cell
	 */
	function add_cell_list($cell) {
		if (array_key_exists($cell, $this->loaded_workbook_cells)) $this->filter[]=$cell;
	}
	
	/**
	 * Print Rows
	 * print sheet cell rows as JSON Document, it will basicly be { "excel cell":"excel data" }
	 */
	function print_sheet_rows() {
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->loaded_workbook_rows);
	}
	
	/**
	 * Print Sheet
	 * print sheet data as JSON Document, it will basicly be { "excel cell" : { "row": "row number", "val": "excel data" } }
	 */
	function print_sheet_data() { 
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->loaded_workbook_data); 
	}
	
	/**
	 * Filter Data
	 * Filter workbook data 
	 * @return array filter data
	 */
	function filter_data() {
		$filter=array();
		foreach ($this->filter as $k => $f)
			if (array_key_exists($f, $this->loaded_workbook_data))
				if ($f != 'sheet')
					if (count($this->filter_label) > 0)
						$filter[$this->filter_label[$k]]=$this->loaded_workbook_data[$f];
					else
						$filter[$f]=$this->loaded_workbook_data[$f];

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
		foreach ($this->filter as $k => $f)
			if (array_key_exists($f, $this->loaded_workbook_cells))
				if (count($this->filter_label) > 0)
					$filter[$this->filter_label[$k]]=$this->loaded_workbook_cells[$f];
				else
					$filter[$f]=$this->loaded_workbook_cells[$f];

		return $filter;
		
	}
	
	/**
	 * Print filtered cells
	 * Print filtered sheet cells as JSON Document, it will basicly be { "excel cell":"excel data" }
	 *
	 */
	function print_filter_cells() {
		if (!$this->loaded_workbook) $this->print_sheets();
		$this->json_out($this->filter_cells());
	}
	
	/**
	 * List Workbooks
	 * print workbook sheet names as JSON Document
	 *
	 * it will basicly be [ "sheet name" ]
	 */
	function print_sheets() { $this->json_out($this->excel_sheets['title']); }
	
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
row="<xsl:value-of select="../@r"/>"
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
