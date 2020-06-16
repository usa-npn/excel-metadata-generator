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
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


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
define("MAGNITUDE_TYPE", 'magnitude');

$ancilliary_types = array(
    'Dataset' => 'dataset', 
    'Person' => 'person', 
    'Site' => 'station',
    'Individual_Plant' => 'plant', 
    'Protocol' => 'protocol', 
    'Species_Protocol' => 'species_protocol', 
    'Phenophase' => 'phenophase', 
    'Phenophase_Definition' => 'phenophase_definition', 
    'Species-Specific_Info' => 'sspi', 
    'Intensity' => 'intensity', 
    'Site_Visit' => 'observation_group'
    );

/**
 * Setup all the various workbooks to generate, including the composite
 * book which will contain all the sheets added to the other books.
 */
$book_composite = createWorkbook("Datafield Descriptions for Status-Intensity and Phenometrics Data", "", "");
$book_composite->createSheet();
$book_composite->createSheet();
$book_composite->createSheet();

$book_raw = createWorkbook("Datafield Descriptions for Status and Intensity Observation Data", "", "");
$book_individual_summarize = createWorkbook("Datafield Descriptions for Individual Phenometrics Data", "", "");
$book_site_summarize = createWorkbook("Datafield Descriptions for Site Phenometrics Data", "", "");
$book_magnitude = createWorkbook("Datafield Descriptions for Magnitude Phenometrics Data", "", "");

$book_ancilliary = createWorkbook("Datafield Descriptions for Ancillary Observation Data", "", "");


addSheet($book_composite, 0, 'Status-Intensity Data', RAW_TYPE);
addSheet($book_raw, 0, 'Status-Intensity Data', RAW_TYPE);

addSheet($book_composite, 1, 'Individual Phenometrics', INDIVIDUAL_TYPE);
addSheet($book_individual_summarize, 0, 'Individual Phenometrics', INDIVIDUAL_TYPE);


addSheet($book_composite, 2, 'Site Phenometrics', SITE_TYPE);
addSheet($book_site_summarize, 0, 'Site Phenometrics', SITE_TYPE);

addSheet($book_composite, 3, 'Magnitude Phenometrics', MAGNITUDE_TYPE);
addSheet($book_magnitude, 0, 'Magnitude Phenometrics', MAGNITUDE_TYPE);

$num_ancilliary_sheets = 0;

foreach($ancilliary_types as $name => $type){
    if($num_ancilliary_sheets > 0){
        $book_ancilliary->createSheet();
    }
    addSheet($book_ancilliary, $num_ancilliary_sheets++, $name, $type);    
}



//Reset the composite metadata sheet to the first page.
$book_composite->setActiveSheetIndex(0);
$book_ancilliary->setActiveSheetIndex(0);

/**
 * Write all the files to disk
 */		
$writer = new Xlsx($book_composite);
$writer->save(OUTPUT_PATH . 'all_datafield_descriptions.xlsx');

$writer = new Xlsx($book_raw);
$writer->save(OUTPUT_PATH . 'status_intensity_datafield_descriptions.xlsx');

$writer = new Xlsx($book_individual_summarize);
$writer->save(OUTPUT_PATH . 'individual_phenometrics_datafield_descriptions.xlsx');

$writer = new Xlsx($book_site_summarize);
$writer->save(OUTPUT_PATH . 'site_phenometrics_datafield_descriptions.xlsx');

$writer = new Xlsx($book_magnitude);
$writer->save(OUTPUT_PATH . 'magnitude_phenometrics_datafield_descriptions.xlsx');

$writer = new Xlsx($book_ancilliary);
$writer->save(OUTPUT_PATH . 'ancillary_datafield_descriptions.xlsx');



/**
 * Returns a PHPExcel object initialized with some basic
 * attributes
 */
function createWorkbook($title, $subject, $description){
    $object = new Spreadsheet();
    $object->getProperties()->setCreator(AUTHOR);
    $object->getProperties()->setLastModifiedBy(AUTHOR);

    $object->getProperties()->setTitle($title);
    $object->getProperties()->setSubject($subject);
    $object->getProperties()->setDescription($description);
    
    return $object;
}

/**
Generic function for adding a sheet.
This is a test comment.
*/
function addSheet(&$object, $sheet_indx, $title, $type){
    $object->setActiveSheetIndex($sheet_indx);
    addHeaders($object);
    
    $object->getActiveSheet()->setTitle($title);
    
    fetchWriteData($type, $object);    
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
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP
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
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM
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
