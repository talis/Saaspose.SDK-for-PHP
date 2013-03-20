<?php

/*
 * reads barcodes from images
 */

class BarcodeReader {

    public $FileName = "";

    public function BarcodeReader($fileName) {
        $this->FileName = $fileName;
    }

    /*
     * reads all or specific barcodes from images
     * @param string $symbology
     */

    public function Read($symbology) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");

            //build URI to read barcode
            $strURI = Product::$BaseProductUri . "/barcode/" . $this->FileName . "/recognize?" .
                    (!isset($symbology) || trim($symbology) === '' ? "type=" : "type=" . $symbology);

            //sign URI
            $signedURI = Utils::Sign($strURI);

            //get response stream
            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            //returns a list of extracted barcodes
            return $json->Barcodes;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function ReadR($remoteImageName, $remoteFolder, $readType) {
        try {
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            $uri = $this->UriBuilder($remoteImageName, $remoteFolder, $readType);
            $signedURI = Utils::Sign($uri);
            $responseStream = Utils::processCommand($signedURI, "GET", "", "");
            $json = json_decode($responseStream);
            return $json->Barcodes;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function ReadFromLocalImage($localImage, $remoteFolder, $barcodeReadType) {
        try {
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            $folder = new Folder();
            $folder->UploadFile($localImage, $remoteFolder);
            $data = $this->ReadR(basename($localImage), $remoteFolder, $barcodeReadType);
            return $data;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
            return null;
        }
    }

    public function UriBuilder($remoteImage, $remoteFolder, $readType) {
        $uri = Product::$BaseProductUri . "/barcode/";
        if ($remoteImage != null && trim($remoteImage) === '')
            $uri .= $remoteImage . "/";
        $uri .= "recognize?";
        if ($readType == "AllSupportedTypes")
            $uri .= "type=";
        else
            $uri .= "type=" . $readType;
        if ($remoteFolder != null && trim($remoteFolder) === '')
            $uri .= "&format=" . $remoteFolder;
        if ($remoteFolder != null && trim($remoteFolder) === '')
            $uri .= "&folder=" . $remoteFolder;
        return $uri;
    }

}
