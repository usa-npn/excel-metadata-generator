<?php
/**
 * 
 * This is a small utility used to fetch metadata from the USA-NPN's webservices
 * and write that info out to a spreadsheet. This is intended to be setup on a
 * cron job or tied to a control that would allow for automatic updates to the 
 * spreadsheets that will be made available to the public after edits are made 
 * on the back-end to the defintions.
 * 
 * 
 * 
 */



/**
 * Setup config params, env. variables, constants
 */
error_reporting(E_ALL);
date_default_timezone_set('America/Phoenix');
$params = parse_ini_file('config.ini');

define('AUTHOR', $params['author']);
define('DOMAIN', $params['domain']);
define('OUTPUT_PATH', $params['output_path']);

define('BASE_ENDPOINT', 'http://' . DOMAIN . '/npn_portal/metadata/getMetadataFields.json');

define('RAW_TYPE', 'raw');
define("INDIVIDUAL_TYPE", 'individual_summarized');
define("SITE_TYPE", 'site_summarized');



include 'Classes/PHPExcel.php';
PHPExcel_Settings::setZipClass(PHPExcel_Settings::PCLZIP);
include 'Classes/PHPExcel/Writer/Excel2007.php';


/**
 * Setup all the various workbooks to generate, including the composite
 * book which will contain all the sheets added to the other books.
 */
$book_composite = createWorkbook("Some Title", "Some Subject", "Some Description");
$book_composite->createSheet();
$book_composite->createSheet();

$book_raw = createWorkbook("Some Title", "Some Subject", "Some Description");
$book_individual_summarize = createWorkbook("Some Title", "Some Subject", "Some Description");
$book_site_summarize = createWorkbook("Some Title", "Some Subject", "Some Description");

addRawSheet($book_composite, 0);
addRawSheet($book_raw, 0);

addIndividualSummarizedSheet($book_composite, 1);
addIndividualSummarizedSheet($book_individual_summarize, 0);

addSiteSummarizedSheet($book_composite, 2);
addSiteSummarizedSheet($book_site_summarize, 0);


//Reset the composite metadata sheet to the first page.
$book_composite->setActiveSheetIndex(0);

/**
 * Write all the files to disk
 */		
$objWriter = new PHPExcel_Writer_Excel2007($book_composite);
$objWriter->save(OUTPUT_PATH . 'composite_metadata.xlsx');

$objWriter = new PHPExcel_Writer_Excel2007($book_raw);
$objWriter->save(OUTPUT_PATH . 'raw_metadata.xlsx');

$objWriter = new PHPExcel_Writer_Excel2007($book_individual_summarize);
$objWriter->save(OUTPUT_PATH . 'individual-summarized_metadata.xlsx');

$objWriter = new PHPExcel_Writer_Excel2007($book_site_summarize);
$objWriter->save(OUTPUT_PATH . 'site-summarized_metadata.xlsx');


/**
 * Returns a PHPExcel object initialized with some basic
 * attributes
 */
function createWorkbook($title, $subject, $description){
    $object = new PHPExcel();
    $object->getProperties()->setCreator(AUTHOR);
    $object->getProperties()->setLastModifiedBy(AUTHOR);

    $object->getProperties()->setTitle($title);
    $object->getProperties()->setSubject($subject);
    $object->getProperties()->setDescription($description);
    
    return $object;
}


/**
 * This grabs the raw metadata type fields and adds them to a sheet.
 * @param PHPExcel $object : The workbook on which to operate
 * @param int $sheet_indx : The index of the work sheet for which to add to the workbook
 * This is necessary because for the composite book, we don't necessarily want to add the 
 * sheet to the beginning of the book.
 */
function addRawSheet(&$object, $sheet_indx){
    $object->setActiveSheetIndex($sheet_indx);
    addHeaders($object);
    $object->getActiveSheet()->setTitle('Raw');    
    
    fetchWriteData(RAW_TYPE, $object);

}

/**
 * This grabs the individual summarized  metadata type fields and adds them to a sheet.
 * @param PHPExcel $object : The workbook on which to operate
 * @param int $sheet_indx : The index of the work sheet for which to add to the workbook.
 * This is necessary because for the composite book, we don't necessarily want to add the 
 * sheet to the beginning of the book.
 */
function addIndividualSummarizedSheet(&$object, $sheet_indx){
    $object->setActiveSheetIndex($sheet_indx);
    addHeaders($object);
    
    $object->getActiveSheet()->setTitle('Individual-Summarized');
    
    fetchWriteData(INDIVIDUAL_TYPE, $object);
}


/**
 * This grabs the site summarized metadata type fields and adds them to a sheet.
 * @param PHPExcel $object : The workbook on which to operate
 * @param int $sheet_indx : The index of the work sheet for which to add to the workbook.
 * This is necessary because for the composite book, we don't necessarily want to add the 
 * sheet to the beginning of the book.
 */
function addSiteSummarizedSheet(&$object, $sheet_indx){
    $object->setActiveSheetIndex($sheet_indx);
    addHeaders($object);
    
    $object->getActiveSheet()->setTitle('Site-Summarized');
    
    fetchWriteData(SITE_TYPE, $object);
    
}

/**
 * Adds the column headers to the worksheet, and adds some styles and dimensions
 * to the sheet as well.
 * @param type The workbook on which to operate. This function assumes that the
 * acitve worksheet has already been selected.
 */
function addHeaders(&$object){
    
    /**
     * This style is for the header.
     */
    $styleArray = array(
        'font'  => array(
            'bold'  => true,
            'color' => array('rgb' => '000000'),
            'size'  => 12,
            'name'  => 'Calibri'
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_TOP
        )            
    );    
    
    
    $headers = array(
            'Sequence #',
            'Field name',
            'Field description',
            'Controlled value choices'
            );
    
    //All the fields in this worksheet should use text wrap
    $object->getDefaultStyle()->getAlignment()->setWrapText(true);
    
    
    /**
     * Actually write the headers to the worksheet, and also set the row height
     */
    for($i =0; $i < count($headers); $i++){
        
        $object->getActiveSheet()->setCellValueByColumnAndRow($i, 1, $headers[$i]);        
        $object->getActiveSheet()->getStyleByColumnAndRow($i, 1)->applyFromArray($styleArray);                
        $object->getActiveSheet()->getRowDimension(1)->setRowHeight(31.5);
    }
    
    /**
     * Each column needs a different amount of space.
     */
    $object->getActiveSheet()->getColumnDimension('A')->setWidth(9.25);
    $object->getActiveSheet()->getColumnDimension('B')->setWidth(27.5);
    $object->getActiveSheet()->getColumnDimension('C')->setWidth(64.38);
    $object->getActiveSheet()->getColumnDimension('D')->setWidth(31.75);
    

}

//column, row
/**
 * This will actually go fetch the metadata from the NPN webservice's and
 * populate the sheet with the information
 * 
 * @param string $type : the metadata type for which to fetch data, e.g. 'raw'
 * 'individual_summary' or 'site_summary'
 * @param PHPExcel $object : The workbook to operate on
 */
function fetchWriteData($type, &$object){
    
    $json = file_get_contents(BASE_ENDPOINT . "?type=" . $type);
    $data = json_decode($json);
    
    
    /**
     * All non-header cells use this style
     */
    $styleArray = array(
        'alignment' => array(
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_BOTTOM
        )            
    );    
    
    for($i=0;$i < count($data); $i++){
        $field = $data[$i];
        //The rows seem to be 1-indexed and we need to skip the header row
        $row_num = $i+2;
        
        $object->getActiveSheet()->getRowDimension($row_num)->setRowHeight(31.5);
        $object->getActiveSheet()->getStyle($row_num)->applyFromArray($styleArray);
        
        $object->getActiveSheet()->setCellValueByColumnAndRow(0, $row_num, $field->seq_num);
        $object->getActiveSheet()->setCellValueByColumnAndRow(1, $row_num, $field->field_name);        
        $object->getActiveSheet()->setCellValueByColumnAndRow(2, $row_num, $field->field_description);
        $object->getActiveSheet()->setCellValueByColumnAndRow(3, $row_num, $field->controlled_values);
    }
    
    //Last step - freeze the sheet's top row for easier browsing
    $object->getActiveSheet()->freezePane('A2');
    
}
