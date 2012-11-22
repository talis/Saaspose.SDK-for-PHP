<?php
/*
* reads barcodes from images
*/
class BarcodeReader
{
	public $FileName = "";
	
    public function BarcodeReader($fileName)
    {
        $this->FileName = $fileName;
    }
	/*
    * reads all or specific barcodes from images
	* @param string $symbology
	*/
	public function Read($symbology, $url=null, $checksumValidation=null, $barcodesCount=0, $folder=null, $storage=null, $rectX=0, $rectY=0, $rectWidth=0, $rectHeight=0)
	{
	    try
		{

            //build URI to read barcode
			$strURI = Product::$BaseProductUri . "/barcode" . (strlen($this->FileName) <= 0 ? "" : "/" . $this->FileName) . "/recognize?" .
						(!isset($symbology) || trim($symbology)==='' ? "type=" : "type=" . $symbology) .
						(strlen($checksumValidation) <= 0 ? "" : "&checksumValidation=" . $checksumValidation) .
						(strlen($url) <= 0 ? "" : "&url=" . $url) .
						($barcodesCount <= 0 ? "" : "&barcodesCount=" . $barcodesCount) .
						(strlen($folder) <= 0 ? "" : "&folder=" . $folder) .
						(strlen($storage) <= 0 ? "" : "&storage=" . $storage) .
						($rectX <= 0 ? "" : "&rectX=" . $rectX) .
						($rectY <= 0 ? "" : "&rectY=" . $rectY) .
						($rectWidth <= 0 ? "" : "&rectWidth=" . $rectWidth) .
						($rectHeight <= 0 ? "" : "&rectHeight=" . $rectHeight); 
			
			if ((strlen($this->FileName) <= 0) AND (strlen($url) > 0) AND ($barcodesCount <= 0) AND (strlen($folder) <= 0) AND (strlen($storage) <= 0) AND ($rectX <= 0) AND ($rectY <= 0) AND ($rectWidth <= 0) AND ($rectHeight <= 0))
			{
				//sign URI
				$signedURI = Utils::Sign($strURI);
				
				//get response stream
				$responseStream = Utils::processCommand($signedURI, "POST", "", "");
			
				$json = json_decode($responseStream);
			
				//returns a list of extracted barcodes
				return $json->Barcodes;
			}
			else
			{
				//check whether file is set or not
				if ($this->FileName == "")
				throw new Exception("No file name specified");
				
				//sign URI
				$signedURI = Utils::Sign($strURI);
				
				//get response stream
				$responseStream = Utils::processCommand($signedURI, "GET", "", "");
			
				$json = json_decode($responseStream);
				
				//returns a list of extracted barcodes
				return $json->Barcodes;
			}

		}
		catch (Exception $e)
		{
			throw new Exception($e->getMessage());
		}  	
	}
}
?>