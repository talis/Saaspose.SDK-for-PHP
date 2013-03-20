<?php

/*
 * This class contains features to work with charts
 */

class CellsWorksheet {

    public $FileName = "";
    public $WorksheetName = "";

    public function CellsWorksheet($fileName, $worksheetName) {
        $this->FileName = $fileName;
        $this->WorksheetName = $worksheetName;
    }

    /*
     * Gets a list of cells
     * $offset
     * $count
     */

    public function GetCellsList($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells?offset=" .
                    $offset . "&count=" . $count;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            $listCells = array();

            foreach ($json->Cells->CellList as $cell) {
                $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                        "/worksheets/" . $this->WorksheetName . "/cells" . $cell->link->Href;

                $signedURI = Utils::Sign($strURI);

                $responseStream = Utils::processCommand($signedURI, "GET", "", "");
                $json = json_decode($responseStream);

                array_push($listCells, $json->Cell);
            }

            return $listCells;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a list of rows from the worksheet
     */

    public function GetRowsList() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/rows";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            $listRows = array();

            foreach ($json->Rows->RowsList as $row) {
                $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                        "/worksheets/" . $this->WorksheetName . "/cells/rows" . $row->link->Href;

                $signedURI = Utils::Sign($strURI);

                $responseStream = Utils::processCommand($signedURI, "GET", "", "");
                $json = json_decode($responseStream);

                array_push($listRows, $json->Row);
            }

            return $listRows;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a list of columns from the worksheet
     */

    public function GetColumnsList() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/columns";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            $listColumns = array();

            foreach ($json->Columns->ColumnsList as $column) {
                $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                        "/worksheets/" . $this->WorksheetName . "/cells/columns" . $column->link->Href;

                $signedURI = Utils::Sign($strURI);

                $responseStream = Utils::processCommand($signedURI, "GET", "", "");
                $json = json_decode($responseStream);

                array_push($listColumns, $json->Column);
            }

            return $listColumns;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets maximum column index of cell which contains data or style
     * $offset
     * $count
     */

    public function GetMaxColumn($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells?offset=" .
                    $offset . "&count=" . $count;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return $json->Cells->MaxColumn;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets maximum row index of cell which contains data or style
     * $offset
     * $count
     */

    public function GetMaxRow($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells?offset=" .
                    $offset . "&count=" . $count;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return $json->Cells->MaxRow;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets cell count in the worksheet 
     * $offset
     * $count
     */

    public function GetCellsCount($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells?offset=" .
                    $offset . "&count=" . $count;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return $json->Cells->CellCount;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets AutoShape count in the worksheet 
     */

    public function GetAutoShapesCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/autoshapes";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return Count($json->AutoShapes->AuotShapeList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specific AutoShape from the sheet
     * $index
     */

    public function GetAutoShapeByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/autoshapes/" . $index;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return $json->AutoShape;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets charts count in the worksheet 
     */

    public function GetChartsCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/charts";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return Count($json->Charts->ChartList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specific chart from the sheet
     * $index
     */

    public function GetChartByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/charts/" . $index;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return $json->Chart;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets hyperlinks count in the worksheet 
     */

    public function GetHyperlinksCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/hyperlinks";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return Count($json->Hyperlinks->HyperlinkList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * Gets a specific hyperlink from the sheet
     * $index
     */

    public function GetHyperlinkByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/hyperlinks/" . $index;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return $json->Hyperlink;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    ///////////new functions ////////////////////

    public function GetComment($cellName) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/comments/" . $cellName;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return $json->Comment->HtmlNote;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetOleObjectByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/oleobjects/" . $index;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return $json->OleObject;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetPictureByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/pictures/" . $index;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);

            return $json->Picture;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetValidationByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/validations/" . $index;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return $json->Validation;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetMergedCellByIndex($index) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/mergedCells/" . $index;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return $json->MergedCell;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetMergedCellsCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/mergedCells";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return $json->MergedCells->Count;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetValidationsCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/validations";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return $json->Validations->Count;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetPicturesCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/pictures";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return count($json->Pictures->PictureList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetOleObjectsCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/oleobjects";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return count($json->OleObjects->OleOjectList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetCommentsCount() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether workshett name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/comments";

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");

            $json = json_decode($responseStream);
            return count($json->Comments->CommentList);
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function HideWorksheet() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/visible?isVisible=false";
            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "PUT", "", "");
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function UnhideWorksheet() {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/visible?isVisible=true";
            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "PUT", "", "");
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function MoveWorksheet($worksheetName, $position) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $fieldsArray["DestinationWorsheet"] = $worksheetName;
            $fieldsArray["Position"] = $position;
            $json = json_encode($fieldsArray);
            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/position";
            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "POST", "json", $json);
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function CalculateFormula($formula) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/formulaResult?formula=" . $formula;
            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");
            $json = json_decode($responseStream);
            return $json->Value->Value;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function SetCellValue($cellName, $valueType, $value) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/" . $cellName . "?value=" . $value . "&type=" . $valueType;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "POST", "", "");
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetRowsCount($offset, $count) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/rows?offset=" . $offset . "&count=" . $count;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");
            $json = json_decode($responseStream);
            return $json->Rows->RowsCount;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetRow($rowIndex) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/rows/" . $rowIndex;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");
            $json = json_decode($responseStream);
            return $json->Row;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function DeleteRow($rowIndex) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/rows/" . $rowIndex;

            $signedURI = Utils::Sign($strURI);
            $responseStream = Utils::processCommand($signedURI, "DELETE", "", "");
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function GetColumn($columnIndex) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/columns/" . $columnIndex;

            $signedURI = Utils::Sign($strURI);

            $responseStream = Utils::processCommand($signedURI, "GET", "", "");
            $json = json_decode($responseStream);
            return $json->Column;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    /*
     * param $dataSort is an array
     * indexes of $dataSort
     * boolean $dataSort["CaseSensitive"]
     * boolean $dataSort["HasHeaders"]
     * int $dataSort["KeyList"]["key"]
     * string $dataSort["KeyList"]["SortOrder"]
     * boolean $dataSort["SortLeftToRight"]
     */

    public function SortData(array $dataSort, $cellArea = "") {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/sort?cellArea=" . $cellArea;
            $json_array = json_encode($dataSort);
            $signedURI = Utils::Sign($strURI);
            $responseStream = Utils::processCommand($signedURI, "POST", "json", $json_array);
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function SetCellStyle($cellName = "", array $style) {
        try {
            //check whether file is set or not
            if ($this->FileName == "")
                throw new Exception("No file name specified");
            //check whether worksheet name is set or not
            if ($this->WorksheetName == "")
                throw new Exception("Worksheet name not specified");

            $strURI = Product::$BaseProductUri . "/cells/" . $this->FileName .
                    "/worksheets/" . $this->WorksheetName . "/cells/" . cellName . "/style";
            $json_array = json_encode($style);
            $signedURI = Utils::Sign($strURI);
            $responseStream = Utils::processCommand($signedURI, "POST", "json", $json_array);
            $json = json_decode($responseStream);
            if ($json->Code == 200)
                return true;
            else
                return false;
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getCell($cellName) {
        try {
            if ($this->FileName == "") {
                throw new Exception("No File Name Specified");
            }
            if ($this->WorksheetName == "") {
                throw new Exception("No Worksheet Specified");
            }
            $str_uri = Product::$BaseProductUri . "/cells/" . $this->FileName . "/worksheets/" . $this->WorksheetName . "/cells/" . $cellName;
            $signed_uri = Utils::Sign($str_uri);
            $response_stream = Utils::processCommand($signed_uri, "GET", "json");
            $json = json_decode($response_stream);
            if ($json->Code == 200) {
                return $json->Cell;
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function getCellStyle($cellName) {
        try {
            if ($this->FileName == "") {
                throw new Exception("No File Name Specified");
            }
            if ($this->WorksheetName == "") {
                throw new Exception("No Worksheet Specified");
            }
            $str_uri = Product::$BaseProductUri . "/cells/" . $this->FileName . "/worksheets/" . $this->WorksheetName . "/cells/" . $cellName . "/style";
            $signed_uri = Utils::Sign($str_uri);
            $response_stream = Utils::processCommand($signed_uri, "GET", "json");
            $json = json_decode($response_stream);
            if ($json->Code == 200) {
                return $json->Style;
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

    public function addPicture($picturePath, $pictureLocation, $upperLeftRow = 0, $upperLeftColumn = 0, $lowerRightRow = 0, $lowerRightColumn = 0) {
        try {
            if ($this->FileName == "") {
                throw new Exception("No File Name Specified");
            }
            if ($this->WorksheetName == "") {
                throw new Exception("No Worksheet Specified");
            }
            if ($pictureLocation == "Server" || $pictureLocation == "server") {
                $str_uri = Product::$BaseProductUri . "/cells/" . $this->FileName . "/worksheets/" . $this->WorksheetName . "/pictures?upperLeftRow=" .
                        $upperLeftRow . "&upperLeftColumn=" . $upperLeftColumn .
                        "&lowerRightRow=" . $lowerRightRow . "&lowerRightColumn=" . $lowerRightColumn .
                        "&picturePath=" . $picturePath;
                $signed_uri = Utils::Sign($str_uri);
                $response_stream = Utils::processCommand($signed_uri, "PUT");
            } else if ($pictureLocation == "Local" || $pictureLocation == "local") {
                if (!file_exists($picturePath)) {
                    throw new Exception("File Does not Exists");
                }
                $stream = file_get_contents($picturePath);
                $str_uri = Product::$BaseProductUri + "/cells/" . $this->FileName . "/worksheets/" . $this->WorksheetName . "/pictures?upperLeftRow=" .
                        $upperLeftRow . "&upperLeftColumn=" . $upperLeftColumn .
                        "&lowerRightRow=" . $lowerRightRow . "&lowerRightColumn=" . $lowerRightColumn .
                        "&picturePath=" . $picturePath;
                $signed_uri = Utils::Sign($str_uri);
                $response_stream = Utils::processCommand($signed_uri, "PUT", "", $stream);
            }
            $json = json_decode($response_stream);
            if ($json->Code == 200) {
                return true;
            } else {
                return false;
            }
        } catch (Exception $e) {
            throw new Exception($e->getMessage());
        }
    }

}