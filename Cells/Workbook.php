<?php
/*
* This class contains features to work with charts
*/
class CellsWorkbook
{
	public $FileName = "";
    public function CellsWorkbook($fileName)
    {
        $this->FileName = $fileName;		
    }		
	
	/*
    * Get Document's properties
	*/		
	public function GetProperties(){
		try{
			//check whether files are set or not
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
						
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName . "/documentProperties";
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "GET", "", "");
			
			$json = json_decode($responseStream);									
						
			
			if($json->Code == 200)
				return $json->DocumentProperties->DocumentPropertyList;
			else
				return false;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Get Resource Properties information like document source format, IsEncrypted, IsSigned and document properties
	@param string $propertyName
	*/		
	public function GetProperty($propertyName){
		try{
			//check whether files are set or not
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			if ($propertyName == "")
				throw new Exception("Property Name not specified");
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName . "/documentProperties/" . $propertyName;
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "GET", "", "");
			
			$json = json_decode($responseStream);
								
			
			if($json->Code == 200)
				return $json->DocumentProperty;
			else
				return false;			
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Set document property
	@param string $propertyName
	@param string $propertyValue
	*/		
	public function SetProperty($propertyName="",$propertyValue=""){
		try{
			//check whether files are set or not
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			if ($propertyName == "")
				throw new Exception("Property Name not specified");
			
			if ($propertyValue == "")
				throw new Exception("Property Value not specified");
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName . "/documentProperties/" . $propertyName;
			
			$put_data_arr['Value'] = $propertyValue;
			
			$put_data = json_encode($put_data_arr);
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "PUT", "json", $put_data);
			
			$json = json_decode($responseStream);											
			
			if($json->Code == 200)
				return $json->DocumentProperty;
			else
				return false;			
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Remove All Document's properties
	*/		
	public function RemoveAllProperties(){
		try{
			//check whether files are set or not
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
						
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName . "/documentProperties";
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "DELETE");
			
			$json = json_decode($responseStream);
						
			
			if(is_object($json))
			{
				if($json->Code == 200)
					return true;
				else
					return false;
			}
			
			return true;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Delete a document property
	@param string $propertyName
	*/		
	public function RemoveProperty($propertyName){
		try{
			//check whether files are set or not
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			if ($propertyName == "")
				throw new Exception("Property Name not specified");
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName . "/documentProperties/" . $propertyName;
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "DELETE", "", "");
			
			$json = json_decode($responseStream);									
			
			if($json->Code == 200)
				return true;
			else
				return false;			
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Create Empty Workbook
	*/		
	public function CreateEmptyWorkbook(){
		try{
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ;
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "PUT");
			
			$json = json_decode($responseStream);
						
			return $json;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Create Empty Workbook
	* @param string $templateFileName
	*/		
	public function CreateWorkbookFromTemplate($templateFileName){
		try{
			
			if ($templateFileName == "")
				throw new Exception("Template file not specified");
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."?templatefile=".$templateFileName;
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "PUT");
			
			$json = json_decode($responseStream);
						
			return $json;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Create Empty Workbook
	* @param string $templateFileName
	* @param string $dataFile	
	*/		
	public function CreateWorkbookFromSmartMarkerTemplate($templateFileName="",$dataFile=""){
		try{
			
			if ($templateFileName == "")
				throw new Exception("Template file not specified");
			
			if ($dataFile == "")
				throw new Exception("Data file not specified");
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."?templatefile=".$templateFileName."&dataFile=".$dataFile;
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "PUT");
			
			$json = json_decode($responseStream);
						
			return $json;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Process Smartmaker Datafile
	* @param string $dataFile	
	*/		
	public function ProcessSmartMarker($dataFile=""){
		try{
			
			if ($templateFileName == "")
				throw new Exception("Template file not specified");
			
			if ($dataFile == "")
				throw new Exception("Data file not specified");
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/smartmarker?xmlFile=".$dataFile;
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "POST");
			
			$json = json_decode($responseStream);
						
			return $json;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Get Worksheets Count in Workbook
	*/		
	public function GetWorksheetsCount(){
		try{
			
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/worksheets";
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "GET");
			$json = json_encode($responseStream);
                        return $json;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Get Names Count in Workbook	
	*/		
	public function GetNamesCount(){
		try{
			
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/names";
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "GET");
			
			$json = json_decode($responseStream);
			return $json->Names.Count;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	/*
    * Get Default Style
	*/		
	public function getDefaultStyle(){
		try{
			
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			
			//build URI to merge Docs
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/defaultStyle";
			
			//sign URI
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "GET");
			
			$json = json_decode($responseStream);
								
			return $json->Names.Count;
					
		} catch (Exception $e) {
			throw new Exception($e->getMessage());
		}
	}
	
	public function EncryptWorkbook($encryptionType="XOR",$password="",$keyLength="")
	{
		try{
			
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			//Build JSON to post
			$fieldsArray["EncriptionType"] = $encryptionType;
			$fieldsArray["KeyLength"] = $keyLength;
			$fieldsArray["Password"] = $password;
			$json = json_encode($fieldsArray);
				

			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/encryption";
			
			$signedURI = Utils::Sign($strURI);

			$responseStream = Utils::processCommand($signedURI, "POST", "json", $json);
				
			$json_response = json_decode($responseStream);

			if($json_response->Code == 200)
				return true;
			else
				return false;			
			
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
		
	}
	
	
	public function ProtectWorkbook($protectionType="all",$password)
	{
		try{
				
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
				
			//Build JSON to post
			$fieldsArray["ProtectionType"] = $protectionType;			
			$fieldsArray["Password"] = $password;
			$json = json_encode($fieldsArray);
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/protection";
				
			$signedURI = Utils::Sign($strURI);
				
			$responseStream = Utils::processCommand($signedURI, "POST", "json", $json);
	
			$json_response = json_decode($responseStream);
						
			
			if($json_response->Code == 200)
				return true;
			else
				return false;
				
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	
	}
	
	public function UnprotectWorkbook($password)
	{
		try{
	
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
	
			//Build JSON to post			
			$fieldsArray["Password"] = $password;
			$json = json_encode($fieldsArray);
	
	
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/protection";
	
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "DELETE", "json", $json);
	
			$json_response = json_decode($responseStream);
						
	
			if($json_response->Code == 200)
				return true;
			else
				return false;
	
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	
	}
	
	public function SetModifyPassword($password)
	{
		try{
	
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
	
			//Build JSON to post
			$fieldsArray["Password"] = $password;
			$json = json_encode($fieldsArray);
	
	
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/writeProtection";
	
			$signedURI = Utils::Sign($strURI);

			$responseStream = Utils::processCommand($signedURI, "POST", "json", $json);
			
			$json_response = json_decode($responseStream);
	
	
			if($json_response->Status == "OK")
				return true;
			else
				return false;
	
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	
	}
	
	public function ClearModifyPassword($password)
	{
		try{
	
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
	
			//Build JSON to post
			$fieldsArray["Password"] = $password;
			$json = json_encode($fieldsArray);
	
	
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/writeProtection";
	
			$signedURI = Utils::Sign($strURI);
	
			$responseStream = Utils::processCommand($signedURI, "DELETE", "json", $json);
	
			$json_response = json_decode($responseStream);
	
	
			if($json_response->Status == "OK")
				return true;
			else
				return false;
	
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	
	}
	
	///////new functions/////////
	
	public function DecryptWorkbook($password){
		try{
			
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			//Build JSON to post
			$fieldsArray["Password"] = $password;
			$json = json_encode($fieldsArray);
				

			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/encryption";
			
			$signedURI = Utils::Sign($strURI);

			$responseStream = Utils::processCommand($signedURI, "DELETE", "json", $json);
				
			$json_response = json_decode($responseStream);
			if($json_response->Code == 200)
				return true;
			else
				return false;			
			
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	}
	public function AddWorksheet($worksheetName){
		try{
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/worksheets/" . $worksheetName;
			
			$signedURI = Utils::Sign($strURI);

			$responseStream = Utils::processCommand($signedURI, "PUT","","");
				
			$json_response = json_decode($responseStream);
			if($json_response->Code == 201)
				return true;
			else
				return false;			
			
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	}
	
	public function RemoveWorksheet($worksheetName){
		try{
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/worksheets/" . $worksheetName;
			
			$signedURI = Utils::Sign($strURI);

			$responseStream = Utils::processCommand($signedURI, "DELETE","","");
				
			$json_response = json_decode($responseStream);
			if($json_response->Code == 200)
				return true;
			else
				return false;			
			
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	}
	public function MergeWorkbook($mergefileName){
		try{
			if ($this->FileName == "")
				throw new Exception("Base file not specified");
			
			$strURI = Product::$BaseProductUri . "/cells/" . $this->FileName ."/merge?mergeWith=" . $mergefileName;
			
			$signedURI = Utils::Sign($strURI);

			$responseStream = Utils::processCommand($signedURI, "POST","","");
				
			$json_response = json_decode($responseStream);
			if($json_response->Code == 200)
				return true;
			else
				return false;			
			
		} catch (Exception $e){
			throw new Exception($e->getMessage());
		}
	}

}
?>