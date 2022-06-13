<?php

require_once 'Classes/PHPExcel.php';

/**
 * Description of ObjectLineExcel
 *
 * @author RemOpt
 */
class ObjectLineExcel {

    public $Vector = array();
	public $objPHPExcel;
	
    function __construct($pathExcel, $extension = 'xls') {
        if($extension == 'xls'){
			$objReader = new PHPExcel_Reader_Excel5();
			$objReader->setReadDataOnly(true);
			$this->objPHPExcel = $objReader->load($pathExcel);
			$this->objPHPExcel->setActiveSheetIndex(0);
		}
		else{
			$objReader = new PHPExcel_IOFactory();
			//$objReader->setReadDataOnly(true);
			$this->objPHPExcel = $objReader->load($pathExcel);
			$this->objPHPExcel->setActiveSheetIndex(0);
		}
    }
	
	public function getAllSheets(){
		return $this->objPHPExcel->getAllSheets();
	}
	public function setActiveSheet($numberName){
		if(is_numeric($numberName)){
			$this->objPHPExcel->setActiveSheetIndex($numberName);
		}else{
			$this->objPHPExcel->setActiveSheetIndexByName($numberName);
		}
	}
	public function getActiveSheet(){
		return $this->objPHPExcel->getActiveSheet();
	}

	public function getSheetByName($pName = '')
    {
        $worksheetCount = count($this->objPHPExcel->_workSheetCollection);
        for ($i = 0; $i < $worksheetCount; ++$i) {
            if ($this->objPHPExcel->_workSheetCollection[$i]->getTitle() === $pName) {
                return $this->objPHPExcel->_workSheetCollection[$i];
            }
        }
    }
	
	public function countColumn(){
		$i =  $this->objPHPExcel->getActiveSheet()->getHighestColumn();
		return PHPExcel_Cell::columnIndexFromString($i);
	}
	
	public function loadRows(){
		$lastRow = $this->objPHPExcel->getActiveSheet()->getHighestRow();
		$lastColumn = $this->objPHPExcel->getActiveSheet()->getHighestColumn();
		$this->Vector = array();
        for ($linha = 1; $linha <= $lastRow; $linha++) {
            $Lines = array();
            for ($coluna = 0; $coluna < PHPExcel_Cell::columnIndexFromString($lastColumn); $coluna++) {
                $Lines[] = $this->objPHPExcel->getActiveSheet()->getCellByColumnAndRow($coluna, $linha)->getValue();
            }
            $this->Vector[] = $Lines;
        }
		return $this->Vector;
	}

	public function getHeaderIndex($cabecalho,$sheetName,$auto = true,$firstCollumn = 'A',$offsetRow = '1'){
		//obs: todos itens do array de cabeçalho devem estar em minúsculo e sem caracteres especiais como ascentos e cedilhas
		if($sheetName!=='' && !is_numeric($sheetName)){
			$sheets = array($this->getSheetByName($sheetName));
		}
		else{
			$sheets = $this->getAllSheets();
		}
		$retorno = array();
		foreach ($sheets as $currentSheet) {

			if($currentSheet->getSheetState() == "visible"){
				$originalSheetName = $currentSheet->getTitle();
				$this->setActiveSheet($originalSheetName);
				$objWorksheet = $this->getActiveSheet();

				if($auto){
					$lastColumn = $objWorksheet->getHighestColumn();
					$array = $objWorksheet->rangeToArray('A1:'.$lastColumn.'1',NULL,TRUE,TRUE);
				}
				else{
					$lastColumn = $objWorksheet->getHighestColumn();
					$array = $objWorksheet->rangeToArray($firstCollumn.$offsetRow.':'.$lastColumn.$offsetRow,NULL,TRUE,TRUE);
				}
				$arrayCabecalho = $array[0];

				//$arrayCabecalhoBakup = $arrayCabecalho;
				$cabecalhoBackup = $cabecalho;
				$cabecalhoAux = $cabecalho;
				for ($i=0; $i < count($arrayCabecalho); $i++) { //remover os ascentos pq da ruim na hora de colocar tudo minusculo mais pra frente no codigo 
					$arrayCabecalho[$i] = preg_replace(array("/(á|à|ã|â|ä)/","/(Á|À|Ã|Â|Ä)/","/(é|è|ê|ë)/","/(É|È|Ê|Ë)/","/(í|ì|î|ï)/","/(Í|Ì|Î|Ï)/","/(ó|ò|õ|ô|ö)/","/(Ó|Ò|Õ|Ô|Ö)/","/(ú|ù|û|ü)/","/(Ú|Ù|Û|Ü)/","/(ñ)/","/(Ñ)/","/(ç)/","/(Ç)/"),explode(" ","a A e E i I o O u U n N c C"),$arrayCabecalho[$i]); //remove ascentos e ç
				}
				$arrayCabecalho=array_flip($arrayCabecalho); //troca keys e values
				$arrayCabecalho = array_fill_keys(array_keys($arrayCabecalho), -1); //seta todos values com -1
				$arrayCabecalho = array_change_key_case($arrayCabecalho, CASE_LOWER); //deixa todas as keys minusculas

				$cabecalhoAux=array_flip($cabecalhoAux); //troca keys e values
				$cabecalhoAux = array_fill_keys(array_keys($cabecalhoAux), -1); //seta todos values com -1
				$cabecalhoAux = array_change_key_case($cabecalhoAux, CASE_LOWER); //deixa todas as keys minusculas

				$foundSomething = false;

				//var_dump($arrayCabecalho);

				foreach ($cabecalho as $campo) {
					$campobackup = $campo;
					$campo = preg_replace(array("/(á|à|ã|â|ä)/","/(Á|À|Ã|Â|Ä)/","/(é|è|ê|ë)/","/(É|È|Ê|Ë)/","/(í|ì|î|ï)/","/(Í|Ì|Î|Ï)/","/(ó|ò|õ|ô|ö)/","/(Ó|Ò|Õ|Ô|Ö)/","/(ú|ù|û|ü)/","/(Ú|Ù|Û|Ü)/","/(ñ)/","/(Ñ)/","/(ç)/","/(Ç)/"),explode(" ","a A e E i I o O u U n N c C"),$campo); //remove ascentos e ç
					$campo = strtolower($campo);
					if(array_key_exists($campo,$arrayCabecalho)){
						$cabecalhoAux[$campobackup]=array_search($campo, array_keys($arrayCabecalho));
						$foundSomething = true;
					}
					else{
						$i = 0;
						foreach ($arrayCabecalho as $nomeColuna => $numeroColuna) {
							$nomeColuna = strtolower($nomeColuna);
							$nomes = explode("|",$campo);
							foreach($nomes as $nome){
								if(strpos($nomeColuna,$nome) !== false){
									$cabecalhoAux[$campobackup]=$i;
									$foundSomething = true;
								}
								// var_dump($nome);
								// echo "<br>";
								// var_dump($campo);
								// echo "<br>";
								// var_dump(strpos($nome,$campo) !== false);
								// echo "<br>";
								// echo "<br>";
							}


							$i++;
						}
					}
				}
				$foundEverything = true;
				foreach($cabecalhoAux as $field=>$value){
					if ($value==-1)
						$foundEverything = false;
				}
				if($foundEverything && $foundSomething){	//verifica através de simples comparação por numero de colunas se a sheet que está sendo avaliada possui as mesmas colunas de interesse sendo buscada, sim é facilmente enganado, mas é improvável
					$retorno["status"]="ok";
					$retorno[$originalSheetName] = array_combine($cabecalhoBackup,$cabecalhoAux); // array_combine para retornar array com chaves originais e não todas minusculas
				}
				else if($foundSomething){
					$retorno["status"]="miss columns: ";
					foreach($cabecalhoAux as $field=>$value){
						if ($value==-1)
							$retorno["status"].=$field.";";
					}
					$retorno["status"] = trim($retorno["status"],"/");
					$retorno[$originalSheetName] = $cabecalhoAux;
				}
				else{
					// $retorno["status"]="no match";
					// $retorno[$originalSheetName] = false;
					throw new Exception('Nenhuma das colunas a importar foi encontrada. O arquivo pode estar fora do padrão esperado.');
				}
				
			}
		}
		if(!empty($retorno)){
			return $retorno;
		}
		else{
			return -1;
		}
	
	
	}

}
?>
