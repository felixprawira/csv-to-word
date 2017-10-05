<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Excel;

class GenerateDocumentController extends Controller
{
    public function __construct()
    {

    }

    public function index()
    {
    	Excel::load('csv/Sample Data.csv', function($reader) 
    	{
    		$results = $reader->get();
    		$csvData = $results->toArray();
    		$csv_value = array();
    		
    		foreach ($csvData as $data) 
    		{
    			foreach ($data as $key => $value) 
    			{
    				if ($key == 'answer')
    				{
    					array_push($csv_value, $value);
    				}
	    		}	
    		}

            $name = basename(__FILE__, '.php');
			$source = public_path()."/{$csv_value[0]}.docx";

			$phpWord = new \PhpOffice\PhpWord\PhpWord();
            $phpWord->setDefaultParagraphStyle(array('spacing' => 250));

			$section = $phpWord->addSection();

            $header = $section->addHeader();
			$header->addImage(public_path().'/img/Sample Logo.png',
                array(
                    'positioning'   => 'absolute',
                    'height'        => 60,
                    'marginTop'     => -70,
                    'wrappingStyle' => 'infront',
                )
            );		
            $header->addText($csv_value[1]." ".$csv_value[2], array(), array('align' => 'right'));
            
            $topLineStyle = array('weight' => 1, 'width' => 760, 'height' => 0);
            $section->addLine($topLineStyle);
            
            $dhTableStyleName = 'Document Header';
            $dhTableStyle = array('borderSize' => 1, 'borderColor' => 'FFFFFF');
            $dhTableFirstRowStyle = array('borderBottomSize' => 1, 'borderBottomColor' => 'FFFFFF', 'bgColor' => 'FFFFFF');
            $dhTableCellStyle = array('valign' => 'center');
            $dhTableCellBtlrStyle = array('valign' => 'center', 'textDirection' => \PhpOffice\PhpWord\Style\Cell::TEXT_DIR_BTLR);
            $dhTableFontStyle = array('bold' => true);
            $phpWord->addTableStyle($dhTableStyleName, $dhTableStyle, $dhTableFirstRowStyle);
            $table = $section->addTable($dhTableStyleName);

            $table->addRow(300);
            $table->addCell(6000, $dhTableCellStyle)->addText('Product Name: '.$csv_value[4], $dhTableFontStyle);
            $table->addCell(4000, $dhTableCellStyle)->addText('Product No.: '.$csv_value[5], $dhTableFontStyle);
            $section->addTextBreak();

            $section->addText('1. SELECTION OF RISK MANAGEMENT STANDARD', array('bold' => true));
            $section->addText('The following standard is applicable to the Risk Management Plan of Axil Scientific Pte. Ltd.: ');
            $section->addText($csv_value[6]);
            $section->addTextBreak();

            $section->addText('2. PURPOSE', array('bold' => true));
            $section->addText($csv_value[7], null, array('align' => 'both'));
            $section->addTextBreak();

            $section->addText('3. RISK MANAGEMENT ACTIVITIES', array('bold' => true));
            $section->addText($csv_value[8], null, array('align' => 'both'));
            $section->addTextBreak();

            $section->addText('SIGNATORY APPROVAL', array('bold' => true));

            $signatureTableStyleName = 'Signatory Table';
            $signatureTableStyle = array('borderSize' => 1, 'borderColor' => '000000');
            $signatureTableFirstRowStyle = array('borderBottomSize' => 1, 'borderBottomColor' => '000000', 'bgColor' => '000000');
            $signatureTableCellStyle = array('valign' => 'center');
            $signatureTableCellBtlrStyle = array('valign' => 'center', 'textDirection' => \PhpOffice\PhpWord\Style\Cell::TEXT_DIR_BTLR);
            $signatureTableFontStyle = array('bold' => true);
            $phpWord->addTableStyle($signatureTableStyleName, $signatureTableStyle, $signatureTableFirstRowStyle);
            $table = $section->addTable($signatureTableStyleName);

            $table->addRow(300);
            $table->addCell(1830, $signatureTableCellStyle)->addText('');
            $table->addCell(1830, $signatureTableCellStyle)->addText('Name', array('bold' => true));
            $table->addCell(1830, $signatureTableCellStyle)->addText('Designation', array('bold' => true));
            $table->addCell(1830, $signatureTableCellStyle)->addText('Signature', array('bold' => true));
            $table->addCell(1830, $signatureTableCellStyle)->addText('Date', array('bold' => true));

            $table->addRow(300);
            $table->addCell(1830, $signatureTableCellStyle)->addText('Prepared by:');
            $table->addCell(1830, $signatureTableCellStyle)->addText($csv_value[9]);
            $table->addCell(1830, $signatureTableCellStyle)->addText($csv_value[10]);
            $table->addCell(1830, $signatureTableCellStyle)->addText('');
            $table->addCell(1830, $signatureTableCellStyle)->addText('');

            $table->addRow(300);
            $table->addCell(1830, $signatureTableCellStyle)->addText('Approved by:');
            $table->addCell(1830, $signatureTableCellStyle)->addText($csv_value[11]);
            $table->addCell(1830, $signatureTableCellStyle)->addText($csv_value[12]);
            $table->addCell(1830, $signatureTableCellStyle)->addText('');
            $table->addCell(1830, $signatureTableCellStyle)->addText('');

            $footer = $section->addFooter();
            $lineStyle = array('weight' => 1, 'width' => 760, 'height' => 0);
            $footer->addLine($lineStyle);
            $footer->addPreserveText('Page {PAGE} of {NUMPAGES}.', null, array('align' => 'right'));
            $footer->addText($csv_value[3], null, array('align' => 'right'));

			$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $file = $csv_value[0].'.docx';
            header("Content-Description: File Transfer");
            header('Content-Disposition: attachment; filename="' . $file . '"');
            header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            header('Content-Transfer-Encoding: binary');
            header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
            header('Expires: 0');
            $xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
            $xmlWriter->save("php://output");
		});
    }
}
