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
        $csv_value = array();
    	$results = Excel::load('csv/Sample Data.csv')->get();
    		
		$csvData = $results->toArray();
		
		foreach ($csvData as $data) 
		{
			foreach ($data as $key => $value) 
			{
				if ($key == 'answer')
				{
					array_push($csv_value, htmlspecialchars($value));
				}
    		}	
		}
        $path = "img/Sample Logo.png";
        $imageSize = getimagesize($path);
        array_push($csv_value, $path);
        array_push($csv_value, $imageSize[0]);
        array_push($csv_value, $imageSize[1]);
        return response()->json($csv_value);
    }
}
