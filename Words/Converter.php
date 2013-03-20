<?php

/*
 * converts pages or document into different formats
 */

class WordConverter {

    public $FileName = "";
    public $saveformat = "";

    public function WordConverter($fileName) {
        //set default values
        $this->FileName = $fileName;

        $this->saveformat = "Doc";
    }

    /*
     * convert a document to SaveFormat
     */

    public function Convert($output_name) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");

            //build URI
            $strURI = Product::$BaseProductUri . "/words/" . $this->FileName . "?format=" . $this->saveformat;

            //sign URI
            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $v_output = Utils::ValidateOutput($responseStream);

            if ($v_output === "") {
                if ($output_name == "") {
                    Utils::saveFile($responseStream, SaasposeApp::$OutPutLocation . Utils::getFileName($this->FileName) . "." . $this->saveformat);
                    return "";
                } else {
                    Utils::saveFile($responseStream, SaasposeApp::$OutPutLocation . $output_name . "." . $this->saveformat);
                    return "";
                }
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function ConvertLocalFile($input_path, $output_path, $output_format) {
        try {
            $str_uri = Product::$BaseProductUri + "/words/convert?format=" + $output_format;
            $signed_uri = Utils::Sign($str_uri);
            $responseStream = Utils::uploadFileBinary($signed_uri, $input_path, "xml");

            $v_output = Utils::ValidateOutput($responseStream);

            if ($v_output === "") {
                if ($output_format == "html")
                    $save_format = "zip";
                else
                    $save_format = $output_format;

                if ($output_path == "") {
                    $outputFilename = Utils::getFileName($input_path) . "." . $save_format;
                }

                Utils::saveFile($responseStream, SaasposeApp::$OutPutLocation . $outputFilename);
                return "";
            }
            else
                return $v_output;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}