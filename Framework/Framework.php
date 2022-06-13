<?php
    require_once 'importExcel/ObjectLineExcel.php';
    require_once 'Connection.php';
    include "PHPMailer/class.phpmailer.php";
    include "PHPMailer/class.pop3.php";
    include "PHPMailer/class.smtp.php";
    include "pdf/tcpdf.php";
	require_once 'dompdf/autoload.inc.php';

    mb_internal_encoding("UTF-8");
    date_default_timezone_set('America/Sao_Paulo');
    setlocale(LC_ALL, '');
	setlocale(LC_NUMERIC, 'en_US.utf8');
	setlocale(LC_TIME, 'pt_BR', 'pt_BR.utf-8', 'pt_BR.utf-8', 'portuguese');
	$path = __DIR__;
	abstract class Utility{
		public static $_userHeaderRqt = "eduardoatene";		


		public static function get_web_page( $url,$json="" )
		{
			$options = array(
				CURLOPT_RETURNTRANSFER => true,     // return web page
				CURLOPT_HEADER         => false,    // don't return headers
				CURLOPT_FOLLOWLOCATION => true,     // follow redirects
				CURLOPT_ENCODING       => "",       // handle all encodings
				CURLOPT_USERAGENT      => "spider", // who am i
				CURLOPT_AUTOREFERER    => true,     // set referer on redirect
				CURLOPT_CONNECTTIMEOUT => 120,      // timeout on connect
				CURLOPT_TIMEOUT        => 120,      // timeout on response
				CURLOPT_MAXREDIRS      => 10,       // stop after 10 redirects
				CURLOPT_SSL_VERIFYPEER => false,     // Disabled SSL Cert checks
				CURLOPT_POST => true,				//use only POST
				CURLOPT_POSTFIELDS => $json,
				CURLOPT_HTTPHEADER => array('Content-Type: application/json; charset=utf-8')
			);
			//curl_setopt($ch, CURLOPT_USERPWD, $username . ":" . $password);  
		
			$ch      = curl_init( $url );
			curl_setopt_array( $ch, $options );
			$content = curl_exec( $ch );
			$err     = curl_errno( $ch );
			$errmsg  = curl_error( $ch );
			$header  = curl_getinfo( $ch );
			curl_close( $ch );
		
			$header['errno']   = $err;
			$header['errmsg']  = $errmsg;
			$header['content'] = $content;
			return $header;
		}

		public static function formatQueryToSearchForColumns($querySql,$columns,$search,$start=-1,$end=-1,$order = array()){
			//echo $querySql;

			$string = "";
			$j=0;
			foreach($search as $word){
				if($word!=""){
					
					if($j>0 && $j<count($search)){
						$string.= " and ";
					}
				
					$c=0;
					$string .="(";
					foreach($columns as $col){								
						if($c>0){
							$string.= " or ";
						}
						$string .= $columns[$c]." like '%".$word."%'";
						$c++;
					}
					$string .=")";
					$j++;
				}
			}
			if($string!="")
				$querySql.=" AND ( ".$string." )";

			if(is_array($order)){		
				if(count($order)>0){
					$str= " ORDER BY ";
					foreach($order as $campo=>$ordem){
						$str.= " ".$campo." ".$ordem.",";
					}
					$str = trim($str,",");
					$querySql.=$str;
				}
			}	
			if(intval($start)>=0 && intval($end)>=0){
				$querySql.=" limit ".$start.",".$end;
			}
			return $querySql;
		}
		
		public static function InsertJavascriptFile($file){
            $path ="./javascript/".$file.".js?rd=".rand();
			//$path ="https://".$_SERVER['HTTP_HOST']."/javascript/".$file.".js?rd=".rand();
			?>
			<script type='text/javascript' language='javascript' src='<?php echo $path ?>'></script>
			<?php
		}

		public static function GetCurrentPage(){
			
			return (Utility::getVariable('status', INPUT_GET)!=null)?Utility::getVariable('status', INPUT_GET):((Utility::getVariable('status', INPUT_POST)!=null)?Utility::getVariable('status', INPUT_POST):"");
		}
		
		public static function dateDiff($data1, $data2, $precision=0){
			if(strstr($data1, "--")!==false){
				$data1 = str_replace("--", "00", $data1);
			}
			$data1 = new DateTime(Utility::dateFormatToEU($data1));
			
			if(strstr($data2, "--")!==false){
				$data2 = str_replace("--", "00", $data2);
			}
			$data2 = new DateTime(Utility::dateFormatToEU($data2));
			$finalDate = "";
			$interval = $data1->diff($data2);
			if($interval!==false){
				if($interval->y>0){
					$finalDate.=$interval->y." ano";
					if($interval->m>1) $finalDate.="s";
					$finalDate.=", ";
				}
				if($interval->m>0){
					$finalDate.=$interval->m." ";
					if($interval->m>1) $finalDate.="meses"; else $finalDate.="mês";
					$finalDate.=", ";
				}
				
				if($interval->d>0){
					$finalDate.=$interval->d." dia";
					if($interval->d>1) $finalDate.="s";
				}
				if($precision==1){
					if($finalDate!="") $finalDate.=", ";
					if($interval->h>0){
						$finalDate.=$interval->h." hora";
						if($interval->h>1) $finalDate.="s";
					}
					if($precision==1){
						if($finalDate!="") $finalDate.=", ";
						if($interval->i>0){
							$finalDate.=$interval->i." minuto";
							if($interval->i>1) $finalDate.="s";
						}
						if($precision==1){
							if($finalDate!="") $finalDate.=", ";
							if($interval->s>0){
								$finalDate.=$interval->s." segundo";
								if($interval->s>1) $finalDate.="s";
							}
						}
					}
				}
				return $finalDate;
			}else{
				return "Error";
			}
		}
		
		public static function dateFormatBRToUS($data){
			if(!is_string($data)) return "00/00/0000";
			if (strstr($data, "/")!==false) {
				$thisDate = Utility::dateFormatToBR($data);
				$thisDate = explode(" ", $thisDate);
				$realDate = $thisDate[0];
				$realDate = explode("/", $realDate);
				$realDate = $realDate[1]."/".$realDate[0]."/".$realDate[2]." ".$thisDate[1];
				
				return $realDate;
			}else{
				return Utility::dateFormat($data);
			}
		}
		public static function dateFormatToBR($data){
			if(!is_string($data)) return "00/00/0000";
			if (strstr($data, "/")===false) {
				return Utility::dateFormat($data);
			}else{
				return $data;
			}
		}
		public static function dateFormatToEU($data){
			if(!is_string($data)) return "0000-00-00";
			if (strstr($data, "-")===false) {
				return Utility::dateFormat($data);
			}else{
				return $data;
			}
		}
		//Função para formatação do campo data de padrão europeu para padrão brasileiro ou vice-versa
		public static function dateFormat($data_solicitada){
			if(!strpos($data_solicitada, " ") ){
				if($data_solicitada==null || trim($data_solicitada)==""){
					return "--/--/----";
				}
				if (strstr($data_solicitada, "/")) {
					$A = explode("/", $data_solicitada);
					if(count($A)>=3){
						$V_data = $A[2] . "-" . $A[1] . "-" . $A[0];
					}else{
						return $data_solicitada;
					}
				} elseif (strstr($data_solicitada, "-")) {
					$A = explode("-", $data_solicitada);
					if(count($A)>=3){
						$V_data = $A[2] . "/" . $A[1] . "/" . $A[0];
					}else{
						return $data_solicitada;
					}
				} else {
					$V_data = "00/00/0000";
				}
				if($V_data=="00/00/0000")
					$V_data = "--/--/----";
				return $V_data;
			}else {
				$data = explode(" ",$data_solicitada);
				return Utility::dateFormat($data[0])." ".Utility::hourFormat($data[1]);
			}
		}

		public static function hourFormat($hour){
			if(strpos($hour, ":")!==false){				
				if($hour=="00:00:00"){
					return "--:--:--";
				}else{
					return $hour;
				}
			}else{
				return $hour;
			}
		}
		public static function compareDates($date1, $date2){
			if(strstr($date1, "/")){
				$date1 = Utility::dateFormat($date1);
			}
			if(strstr($date2, "/")){
				$date2 = Utility::dateFormat($date2);
			}
			if(strtotime($date1)<strtotime($date2)){
				return 1;
			}else if(strtotime($date1)>strtotime($date2)){
				return -1;
			}else{
				return 0;
			}
		}
		
		public static function getVariable($variable, $type=INPUT_POST, $filter=FILTER_SANITIZE_STRING){
			//echo "POST: ".INPUT_POST, "<BR>";
			//echo "GET: ".INPUT_GET, "<BR>";
			//echo $type, " - ", $variable, " - ", filter_input($type, $variable, FILTER_SANITIZE_STRING, FILTER_REQUIRE_ARRAY), "<BR>";
			if(($type==INPUT_POST && isset($_POST[$variable]) && is_array($_POST[$variable])) || ($type==INPUT_POST && isset($_POST[$variable]) && is_array($_POST[$variable]))){
				$options = FILTER_REQUIRE_ARRAY;
			}else{
				$options = FILTER_REQUIRE_SCALAR;
			}
			return filter_input($type, $variable, $filter, $options);
				
			//old
			
			if((isset($_POST[$variable]) && is_array($_POST[$variable])) || 
				(isset($_GET[$variable]) && is_array($_GET[$variable]))){
				$get=filter_input(INPUT_GET, $variable, FILTER_SANITIZE_STRING, FILTER_REQUIRE_ARRAY);
				$post=filter_input(INPUT_POST, $variable, FILTER_SANITIZE_STRING, FILTER_REQUIRE_ARRAY);
			}else{
				$get=filter_input(INPUT_GET, $variable, FILTER_SANITIZE_STRING);
				$post=filter_input(INPUT_POST, $variable, FILTER_SANITIZE_STRING);
			}
			if(isset($get)){
				return $get;
			}else if(isset($post)){
				return $post;
			}else{
				return null;
			}
		}

		public static function doExist($variable){
			if(isset($variable) && trim($variable)!="" && trim($variable)!=0){
				return true;
			}else return false;
		}
		
		function printVarName($var) {
			foreach($GLOBALS as $var_name => $value) {
				if ($value === $var) {
					return $var_name;
				}
			}
			return false;
		}
		function isDate($value){
			if(strpos($value, ":")!==false){
				$value = explode(" ", $value);
				$value = $value[0];
			}
			$values = preg_match('/^(?:(?:31(\/|-|\.)(?:0?[13578]|1[02]))\1|(?:(?:29|30)(\/|-|\.)(?:0?[1,3-9]|1[0-2])\2))(?:(?:1[6-9]|[2-9]\d)?\d{2})$|^(?:29(\/|-|\.)0?2\3(?:(?:(?:1[6-9]|[2-9]\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\d|2[0-8])(\/|-|\.)(?:(?:0?[1-9])|(?:1[0-2]))\4(?:(?:1[6-9]|[2-9]\d)?\d{2})$/', $value);
			return $values;
		}
		function isDouble($value){
			if(is_numeric($value)){
				if(strstr($value, ".") && !strstr($value, ",")){
					return true;
				}else{
					return false;
				}					
			}else{
				return false;
			}
		}
		static function valueToDouble($value){		
			if(trim($value)==""){
				$value="0,00";
			}		
			if(!strstr($value, ",")){
				//return $value;
				$value = Utility::doubleToValue($value);
			}
			$value = str_replace(".", "", $value);
			$value = str_replace(",", ".", $value);
			return $value;
		}
		static function doubleToValue($value){
			if(trim($value)==""){
				$value="0.00";
			}
			if(!is_numeric(str_replace(".", "", str_replace(",", "", $value)))){
				$value="0.00";
			}
			
			
			if((strpos($value, ",")!==false && strpos($value, ".")!==false) || strpos($value, ",")!==false){
				$value= Utility::valueToDouble($value);
			}
			
			$value = number_format($value, 2, ",", ".") or die ("Error: -".$value."-");
			return $value;
		}
		
		public static function ShowMessage($message){ ?>
			<script>
				$(document).ready(function() {
					alert("<?php echo $message ?>");
				});
			</script>
			<?php
		}
		
		public static function ShowTimer($seconds=1500, $success=true, $formName="", $alertType="info"){?>
			<script>
				$(document).ready(function() {
					var action, alerta;
					var type = "<?php echo $alertType ?>";
					<?php
					if($formName==""){ ?>
						var url = location.protocol + '//' + location.host + location.pathname;
						url += "?pagina=<?php echo Utility::getVariable('pagina', INPUT_GET, FILTER_REQUIRE_SCALAR); ?>";
						url = getLinkBySystem(url);
						console.log(url);
						action = function() {
							window.location = url;
						}
					<?php
					}else{  ?>
						var form = $("#<?php echo $formName?>");
						form.attr("action", getLinkBySystem(form.attr("action")));
						action = function() {
							form.submit();
						}
					<?php 
					}
					if(is_bool($success)){
						if($success){?>
							alerta = "Salvo com sucesso.";
					<?php }else{?>
							alerta = "Ocorreu um erro. Contate o suporte do sistema.";
					<?php }
					}else{ 
						if(trim($success)!=""){?>
							alerta = "<?php echo $success ?>";<?php
						}
					}?>
					if(alerta!="" && alerta!=undefined){
						alert(alerta, type);
					}
					setTimeout(action, <?php echo $seconds ?>);
				});
			</script>
			<?php
		}

		public static function validaCPF($cpf) {
 
			// Extrai somente os números
			$cpf = preg_replace( '/[^0-9]/is', '', $cpf );
			 
			// Verifica se foi informado todos os digitos corretamente
			if (strlen($cpf) != 11) {
				return false;
			}
		
			// Verifica se foi informada uma sequência de digitos repetidos. Ex: 111.111.111-11
			if (preg_match('/(\d)\1{10}/', $cpf)) {
				return false;
			}
		
			// Faz o calculo para validar o CPF
			for ($t = 9; $t < 11; $t++) {
				for ($d = 0, $c = 0; $c < $t; $c++) {
					$d += $cpf[$c] * (($t + 1) - $c);
				}
				$d = ((10 * $d) % 11) % 10;
				if ($cpf[$c] != $d) {
					return false;
				}
			}
			return true;
		
		}
	}

	abstract class Element{
		public static function IsVariableEmpty($variable, &$outVar, $output, $debug=false){
			$isEmpty = false;
			if(strpos($variable, "data")!==false){		
				if($variable=="00/00/0000"){
					$isEmpty = true;
				}
			}
			if($variable=="" || $variable==-1){
				$isEmpty = true;
			}
			if($debug){ var_dump( $variable);echo "<BR>"; }
			
			if($isEmpty){
				$outVar.=$output;
				return true;
			}
			return false;
		}
		protected function fillObjectWithQuery($id, $Connection=null){
			$myClass = get_class($this);
			if($Connection==null){
				Connection::SetTable($myClass::$database);
				Connection::SetConditions(array("id"=>$id));
				$fetch = Connection::consult();
			}else{
				$Connection->SetTable($myClass::$database);
				$Connection->SetConditions(array("id"=>$id));
				$fetch = $Connection->consult();
			}
			$vars = $this->getAllLocalVariables();
			$objectVars = $this->getObjectVariables();
			if(($Connection==null && Connection::$numRows>0) || ($Connection!=null && $Connection->numRows>0)){
				$fetch = $fetch[0];
				foreach($vars as $var){
					if (DateTime::createFromFormat('Y-m-d G:i:s', $fetch[$var]) !== FALSE || 
						DateTime::createFromFormat('Y-m-d', $fetch[$var]) !== FALSE) {
						$this->$var = Utility::dateFormatToBR($fetch[$var]);
					}else if(Utility::isDouble($fetch[$var])){
						//$this->$var = Utility::doubleToValue($fetch[$var]);
						$this->$var = $fetch[$var];
					}else{
						$this->$var = $fetch[$var];
					}
				}
				
				/*
				Se ID == null
					OBJ == novo(-1)
				Se ID é um numero
					OBJ = novo(ID)				
				*/
				foreach($objectVars as $var){				
					$nomeClasse = ltrim($var, '_');
					$nomeVariavel = "id".$nomeClasse;
					if(class_exists($nomeClasse)){
						if(isset($this->$nomeVariavel)){
							$this->$var = new $nomeClasse($this->$nomeVariavel);
						}else{
							$nomeObj = "_".$nomeClasse;
							if(!is_object($this->$nomeObj) || $this->$nomeObj==null){
								$this->$nomeObj = new $nomeClasse(-1);
							}
							//echo "não existe: ".$nomeVariavel." em ".get_class($this)."<BR><BR><BR>";
						}
					}else{
						//echo "não existe: ".$nomeClasse."<BR><BR><BR>";
					}
				}
			}else{	
				foreach($vars as $var){
					if(strtoupper($var) == "ID" && $id == -1){
						$this->$var = -1;
					
					}elseif (DateTime::createFromFormat('Y-m-d G:i:s', $var) !== FALSE || 
						DateTime::createFromFormat('Y-m-d', $var) !== FALSE) {
						$this->$var = "00/00/0000";
					}else if(Utility::isDouble($var)){
						$this->$var = "0,00";
					}
				}				
				foreach($objectVars as $var){
					$nomeClasse = ltrim($var, '_');
					$nomeVariavel = "id".$nomeClasse;
					
					if(class_exists($nomeClasse)){
						$this->$var = new $nomeClasse(-1);
					}
				}
			}
		}
		protected function getObjectVariables(){
			$vars = get_object_vars($this);
			$endVariables = array();
			foreach($vars as $key=>$variable){
				if(substr($key, 0, 1) == "_"){
					array_push($endVariables, $key);
				}
			}
			return $endVariables;
		}
		
		protected function getAllLocalVariables(){
			$vars = get_object_vars($this);
			$endVariables = array();
			foreach($vars as $key=>$variable){
				if(substr($key, 0, 1) != "_"){
					array_push($endVariables, $key);
				}
			}
			return $endVariables;
		}
		
		function newInstanceWithoutConstructor($class){
			return unserialize(
				sprintf(
				'O:%d:"%s":0:{}',
				strlen($class), $class
				)
			);
		}
		static function PopulateEmpty($class){
			$ref = new ReflectionClass($class);
			if(method_exists($ref, "newInstanceWithoutConstructor")){
				$object = $ref->newInstanceWithoutConstructor();
			}else{
				$object = Element::newInstanceWithoutConstructor($class);
			}
			$variables = get_object_vars($object);

			foreach($variables as $key=>$attr){
				if(substr($key, 0, 1) == "_"){
					$nomeClasse = $nomeClasse = ltrim($key, '_');
					if(class_exists($nomeClasse)){
						$object->$key = new $nomeClasse(-1);
					}
				}else{
					$object->$key = -1;
				}
			}
			return $object;
		}
		public static function NewInstance($id=-1, $classOverride="-1", $Connection=null){
			$originalClass = get_called_class();
			if($id!=-1){
				$originalClass = Element::PopulateEmpty($originalClass);
				
				if($classOverride!=-1){
					$thisClass = $classOverride;
				}else{
					$thisClass = $originalClass;
				}
				
				$classToCheck = new $thisClass("");
				
				if($Connection==null){
					if(isset($classToCheck::$database)){
						Connection::setTable($classToCheck::$database);
					}else{
						throw new Exception('Class '.$thisClass.' doesn\'t possess a static attribute $database, which is required.');
						die();
					}
					
					Connection::setConditions(array("id"=>$id));
					$variables = get_object_vars($originalClass);
					
					$fetch = Connection::consult();			
					if(Connection::$numRows>0){
						$fetch = $fetch[0];
						
						foreach($variables as $attr){
							$originalClass->$$attr = $fetch[attr];
						}
					}else{
						foreach($variables as $attr){
							$originalClass->$$attr = -1;
						}
					}
				}else{
					if(isset($classToCheck::$database)){
						$Connection->setTable($classToCheck::$database);
					}else{
						throw new Exception('Class '.$thisClass.' doesn\'t possess a static attribute $database, which is required.');
						die();
					}
					
					$Connection->setConditions(array("id"=>$id));
					$variables = get_object_vars($originalClass);
					
					$fetch = $Connection->consult();			
					if($Connection->$numRows>0){
						$fetch = $fetch[0];
						
						foreach($variables as $attr){
							$originalClass->$$attr = $fetch[attr];
						}
					}else{
						foreach($variables as $attr){
							$originalClass->$$attr = -1;
						}
					}
				}
			}else{
				return Element::PopulateEmpty($originalClass);
			}
		}
		public static function GetAll(){
			$Connection = null;
			if (func_num_args())
				$Connection = func_get_arg(0);
			
			$class = get_called_class();
			$className = $class;
			$class = new $class(-1);
			$ativo = "";
			
			if($Connection==null || ($Connection!=null && !is_object($Connection))){
				Connection::setTable($class::$database);
				if(property_exists($class, "nome")){
					$var = "nome";
				}elseif(property_exists($class, "id")){
					$var = "id";
				}
				if(property_exists($class, "ativo")){
					$ativo = "ativo = 1";
					Connection::setConditions($ativo);
				}
				Connection::setOrder("$var asc");
				$fetch=Connection::consult();
			}else{
				$Connection->setTable($class::$database);
				if(property_exists($class, "nome")){
					$var = "nome";
				}elseif(property_exists($class, "id")){
					$var = "id";
				}
				if(property_exists($class, "ativo")){
					$ativo = "ativo = 1";
					Connection::setConditions($ativo);
				}
				$Connection->setOrder("$var asc");
				$fetch=$Connection->consult();
			}
			$new=array();


			foreach($fetch as $key=> $atual){
				if(!isset($atual['id'])){
					$id = $atual["id".get_called_class()];
				}else{
					$id = $atual['id'];
				}
				
				if(isset($new[$atual[$var]])){
					$novoNome = $atual[$var];
					//if($className=="Cidade"){
						$c=0;
						while(isset($new[$novoNome]) && $c<10){
							//echo "já existe uma chave: '".$atual[$var]."', tentando: ";
							$novoNome = $novoNome." ";
							//echo "'".$novoNome."'";
							$c++;
						}		
						//echo "FOI!";
					//}
					

					$new[$novoNome]=$id;
					//echo "<BR>";
				}else{
					$new[$atual[$var]]=$id;
				}
			}
			return $new;
		}

		public function Save(){
			$Connection = null;
			if (func_num_args())
				$Connection = func_get_arg(0);
			
			$idvar = "id";
			if(!isset($this->id)){
				$localVars = $this->getAllLocalVariables();
				$idvar =  $localVars[0];
			}
			$vars = array_values($this->getAllLocalVariables());
			$varsObj = array_values($this->getObjectVariables());

			foreach($varsObj as $nomeVariavel){
				$idReferenciaObjeto = "id".ltrim($nomeVariavel, "_");
				$nomeVariavelObjeto = $nomeVariavel;
				$nomeClasseObjeto = trim($nomeVariavel, "_");
				
				if(/*class_exists($nomeClasseObjeto) && */property_exists($this, $idReferenciaObjeto)/* && !empty($this->$idReferenciaObjeto)*/){
					if($this->$nomeVariavel==null && $this->$idReferenciaObjeto==null) continue;
					
					//Verificar se o objeto foi inicializado
					if(!($this->$nomeVariavelObjeto!=null && is_object($this->$nomeVariavelObjeto))){
						if($this->$idReferenciaObjeto!=null && $this->$idReferenciaObjeto>0 && $this->$idReferenciaObjeto!="" && class_exists($nomeClasseObjeto)){
							$this->$nomeVariavelObjeto = new $nomeClasseObjeto($this->$idReferenciaObjeto);
							throw new Exception("Erro.<br>Classe: ".get_class($this)."<br>Objeto ".$nomeVariavel." foi não foi inicializado, mas possui a variavel ".$idReferenciaObjeto." possui um ID.");
						}
					}else{
					}
					$ID = null;
					if($this->$nomeVariavelObjeto!=null && $this->$nomeVariavelObjeto->isPartiallyFilled()){
						$ID = $this->$nomeVariavelObjeto->Save($Connection);

					}

					if($this->$idReferenciaObjeto==null || $this->$idReferenciaObjeto<=0 || $this->$idReferenciaObjeto==""){
						if($ID!=null && $ID>0){
							$this->$idReferenciaObjeto = $ID;
						}					
					}
				}else{					
				}
			}
			if($this->$idvar==-1 || trim($this->$idvar)==""){
				$name = get_class($this);
				$query = "insert into ".$this::$database." (";
				foreach($vars as $value){
					if($this->$value!="" && ((is_numeric($this->$value) && $this->$value!=-1)||(!is_numeric($this->$value)))){
						$query.="`".$value."`,";
					}
				}
				$query = rtrim($query, ",");
				$query.=") values (";
				foreach($vars as $value){
					if($this->$value!="" && ((is_numeric($this->$value) && $this->$value!=-1)||(!is_numeric($this->$value)))){
						if(Utility::isDate($this->$value)){
							$this->$value = Utility::dateFormatToEU($this->$value);
						}
						$query.="'".$this->$value."',";
					}
				}
				$query = rtrim($query, ",");
				$query.=")";
				if($Connection==null){
					Connection::query($query);
					$this->$idvar = Connection::$lastInsertedId;
				}else{
					$Connection->query($query);
					$this->$idvar = $Connection->lastInsertedId;
				}
			}else{	
				$query = "update ".$this::$database." set ";
				foreach($vars as $value){
					if(($this->$value!="" || $this->$value == null ) &&  ((is_numeric($this->$value) && $this->$value!=-1)||(!is_numeric($this->$value)))){
						if(Utility::isDate($this->$value)){
							$this->$value = Utility::dateFormatToEU($this->$value);
						}
						//parte nova do Bruno Henrique
						// else if(is_numeric(str_replace(".","",str_replace(",",".",$this->$value)))){
						// 	$this->$value = Utility::valueToDouble($this->$value);
						// } 
						//*****************************
						$query.="`".$value."`='".$this->$value."',";
					}
				}
				$query = rtrim($query, ",");
				$query.=" where id='".$this->$idvar."'";
				if($Connection==null){
					Connection::query($query);
				}else{
					$Connection->query($query);
				}
			}
			
			return $this->$idvar;
		}
		public function isPartiallyFilled(){
			$vars = $this->getAllLocalVariables();
			$derp = "";
			//Não contar com a variavel "id"
			array_shift($vars);
			foreach($vars as $variable){
				if($this->IsVariableEmpty($variable, $derp, "")){
					return false;
				}
			}
			return true;
		}
		public function Delete(){
			$Connection = null;
			if (func_num_args())
				$Connection = func_get_arg(0);
			
			$idvar = "id";
			if(!isset($this->id)){
				$derp = $this->getAllLocalVariables();
				$idvar = key(reset($derp[0]));
			}
			if($this->$idvar!=-1){
				$query = "delete from ".$this::$database." where id='".$this->$idvar."'";			
				if($Connection==null){
					Connection::query($query);
				}else{
					$Connection->query($query);
				}				
			}
			return $this->$idvar;
		}
	}

	abstract class DatabaseIteraction extends Element{

		public static $database = "";
		public static function setDatabase($str){
			self::$database = $str;
		}
		public static function getDatabase(){
			return self::$database;
		}
		public static function get(int $id=null){
            $retorno = array("status"=>"","data"=>array());
            $query = "SELECT * FROM ".get_called_class()::$database;
            if($id!=null){
                $query.=" WHERE id = ".$id;
            }
            $fetch=Connection::query($query);
            if(Connection::$numRows>0){
                $retorno["status"]="200";	
                while($rs=$fetch->fetch_array(MYSQLI_ASSOC)){
                    array_push($retorno["data"],$rs);
                }
            }
            else{
                $retorno["status"]="204";	
            }
            return json_encode($retorno);
        }

        public static function search($searchParameters,$caseSensitive=true,$typeSearch="like"){
            $retorno = array("status"=>"","data"=>array());
            $searchParameters = json_decode($searchParameters);
            $query = "SELECT id FROM ".get_called_class()::$database." WHERE";
            $str = "";
            foreach($searchParameters as $field=>$value){
                switch($typeSearch){
                    case "like":
                        if($caseSensitive)
                            $str .= " ".$field." LIKE '%".$value."%' AND";
                        else
                            $str .= " ".strtolower($field)." LIKE LOWER('%".$value."%') AND";
                    break;
                    default: //equals
                        if($caseSensitive)
                            $str .= " ".$field." = ".$value." AND";
                        else
                            $str .= " ".strtolower($field)." = LOWER(".$value.") AND";
                    break;
                }

            }
            $str = trim($str,"AND");
            $query = $query.$str;		
            $fetch=Connection::query($query);
            if(Connection::$numRows>0){
                $retorno["status"]="200";			
                while($rs=$fetch->fetch_array(MYSQLI_ASSOC)){
                    array_push($retorno["data"],$rs["id"]);
                }
            }
            else{
                $retorno["status"]="204";	
            }
            return json_encode($retorno);
        }

        public static function set($data){
            $data = json_decode($data);

			$query = "insert into ".get_called_class()::$database." (";
			foreach($data as $field=>$value){
				//if($value!="" && ((is_numeric($field) && $field!=-1)||(!is_numeric($field)))){
					$query.="`".$field."`,";
				//}
			}
			$query = rtrim($query, ",");
			$query.=") values (";
			foreach($data as $field=>$value){
				if($value!="" && ((is_numeric($value) && $value!=-1)||(!is_numeric($value)))){
					if(Utility::isDate($value)){
						$value = Utility::dateFormatToEU($value);
					}
				}
				$query.="'".$value."',";
			}
			$query = rtrim($query, ",");
			$query.=")";
			Connection::query($query);	
			return Connection::$lastInsertedId;		
        }

        public static function update(int $id,$data){
            $data = json_decode($data);

			$query = "update ".get_called_class()::$database." set ";
					foreach($data as $field=>$value){
						// if(($value!="" || $value == null ) &&  ((is_numeric($value) && $value!=-1)||(!is_numeric($value)))){
						// 	if(Utility::isDate($value)){
						// 		$value = Utility::dateFormatToEU($value);
						// 	}
						// }
						$query.="`".$field."`='".$value."',";
					}
			$query = rtrim($query, ",");
			$query.=" where id='".$id."'";		
			
			Connection::query($query);
        }
	}

	abstract class formElement{
		public $uid, $name, $label, $title, $placeHolderText, $helpText, $value, $active, $required, $visible, $showLabel, $showContainerDiv, $formGroupClass, $formGroupStyle, $style, $class, $rightIcon, $leftIcon;
		public $atributes=array(), $innerElement = array();
		protected $atributesString="";
		public function __construct($uid="-1", $label="Input Element", $helpText="", $placeHolderText=""){
			if(isset($uid)) $this->uid=$uid;
			else $this->uid=-1;
			if(isset($label)) $this->label=ucwords($label);
			else $this->label=-1;
			if(isset($helpText)) $this->helpText=$helpText;
			else $this->helpText=-1;
			if(isset($placeHolderText)) $this->placeHolderText=$placeHolderText;
			else $this->placeHolderText=-1;
			$this->value="";
			$this->active=true;
			$this->required=true;
			$this->visible=true;
			$this->showLabel=true;
			$this->showContainerDiv=true;
			$this->formGroupClass="";
			$this->formGroupStyle="";
			$this->title="";
			$this->style="";
			$this->class="";
			$this->leftIcon="";
			$this->rightIcon="";
			$this->atributes=array();
			$this->labelStyle = "";
			$this->labelClass = "";
		}	
		public function showInnerElement(){
			if(is_array($this->innerElement)){
				foreach($this->innerElement as $row){					
					if(method_exists($row, "Show")){
						$row->Show();
					}else if(method_exists($row, "InnerShow")){
						$row->InnerShow();
					}else{
						echo $row;
					}
				}
			}else{				
				if(method_exists($this->innerElement, "Show")){
					$this->innerElement->Show();
				}else if(method_exists($this->innerElement, "InnerShow")){
					$this->innerElement->InnerShow();
				}else{
					echo $this->innerElement;
				}					
			}
		}
		public function AddAttribute($attr){
			if(is_array($attr)){
				array_push($this->atributes, $attr);
			}else{
				$derp[$attr] = $attr;
				array_push($this->atributes, $derp);
			}
			
			$attrs = "";
			foreach($this->atributes as $atribute){
				foreach($atribute as $key=>$derp){
					$attrs.=" ".$key."=\"".$derp."\"";
				}
			}
			$this->atributesString = $attrs;
		}		
		public function setLabelClass($class){
			$this->labelClass = $class;
		}
		public function setLabelStyle($style){
			$this->labelStyle = $style;
		}
		public function setRightIcon($value){
			$this->rightIcon = $value;
		}
		public function setLeftIcon($value){
			$this->leftIcon = $value;
		}
		public function setClass($value){
			$this->class = $value;
		}
		public function setStyle($name){
			$this->style=$name;
		}

		public function setTitle($name){
			$this->title=$name;
		}
		
		public function setInnerElement($element){
			$this->innerElement=$element;
		}
		
		public function setRequired($bool){
			$this->required=$bool;
		}
		
		public function setName($name){
			$this->name=$name;
		}

		public function setID($id){
			$this->uid=$id;
		}

		public function setVisible($bool){
			$this->visible=$bool;
		}

		public function setActive($bool){
			$this->active=$bool;
		}

		public function setLabel($label){
			$this->label=$label;
		}
		public function setValue($values){
			$this->value=$values;
		}
		public function setLabelVisibility($value){
			$this->showLabel=$value;
		}
		public function showContainerDiv($value){
			$this->showContainerDiv=$value;
		}
		public function setFormGroupClass($value){
			$this->formGroupClass=$value;
		}
		public function setFormGroupStyle($value){
			$this->formGroupStyle=$value;
		}

		public function InsertRow($classRow, $elements){
			array_push($this->rows, new RowElement($classRow, $elements));
		}
		function startElement(){	
			if($this->showContainerDiv){ ?>
				<div style="<?php echo $this->formGroupStyle?>" class="form-group <?php echo $this->formGroupClass ?> <?php if($this->rightIcon!="" || $this->leftIcon){ echo "input-group"; } ?>" <?php if($this->visible!=false) echo "style='display: none'"?>>
				<?php
					if($this->leftIcon!=""){
						echo "<span class=\"input-group-addon\">";
						if(gettype($this->leftIcon)!="string"){
							if(get_parent_class($this->leftIcon)=="formElement"){
								$this->leftIcon->InnerShow();
							}
						}else{
							echo $this->leftIcon;
						}
						echo "</span>";
					}	
			}
		}

		function endElement(){
			if($this->showContainerDiv){
				if($this->rightIcon!=""){ 
					echo "<span class=\"input-group-addon\">";
						if(gettype($this->rightIcon)!="string"){
							if(get_parent_class($this->rightIcon)=="formElement"){
								$this->rightIcon->InnerShow();
							}
						}else{
							echo $this->rightIcon;
						}
					echo "</span>";
				}?>
				</div>
				<?php
			}
		}

		function showLabel(){
			if($this->label!=null){
				?>
				<label style="<?php echo $this->labelStyle ?>" class="control-label <?php echo $this->labelClass ?>"><?php if(trim($this->label)!="" && $this->showLabel){ echo $this->label; }else{ echo "&nbsp;"; } ?></label>
				<?php
			}
		}

		function showHelpText(){
			if($this->helpText){
				?>
				<p class="help-block"> <?php echo $this->helpText; ?></p>
				<?php
			}
		}

		abstract protected function InnerShow();

		public function Show(){
			$this->startElement();
			$this->showLabel();
			$this->InnerShow();
			$this->showHelpText();
			$this->endElement();
		}

	}

	class itemMenu{

		public $name, $icon, $link, $secondList, $thirdList, $color, $newLink, $title;

		public function __construct($name="Sem nome", $icon="", $link="#", $title = ""){
			$this->name = ucwords($name);
			$this->icon = $icon;
			$this->link = $link;
			$this->color = "";
			$this->newLink = "";
			$this->title = $title;
		}
		public function setColor($color){
			$this->color = $color;
		}
		public function setLink($zelda){
			$this->newLink = $zelda."?";
		}
		public function show(){
			$hasList=0;
			if(is_array($this->secondList)){
				if(count($this->secondList)>0) $hasList=2;
			}
			if(is_array($this->thirdList)){
				if(count($this->thirdList)>0) $hasList=3;
			}
			
			$folder = explode("/", $_SERVER["REQUEST_URI"]);
			$_SystemRootRedirectName = $folder[1];			
			?>
			<li style="<?php if($this->color!="") echo "background-color: ".$this->color.";" ?>" class = "nav-item">
				<a title="<?php echo $this->title; ?>" href="<?php echo ($this->newLink == "" ?( "../".$_SystemRootRedirectName."/".(($hasList>0) ? "#" : "?pagina=".$this->link)):$this->newLink) ?>" class = "nav-link">
					<i class="mdi mdi-<?php echo $this->icon?> menu-icon"></i>
					<span><?php echo $this->name;?></span>
					<i class="menu-arrow"></i>
				</a>
			<?php
			if($hasList>0){
				?>
					<div class="submenu">
						<ul class="submenu-item">
				<?php
				foreach(($hasList==2) ? $this->secondList : $this->thirdList as $atual){
					$atual->show();
				}
				?>
						</ul>
					</div>
				<?php }
			?>
			</li>
			<?php
		}

	}
	
	class ButtonElement extends formElement{
		public $maxChars, $name, $icon, $iconStyle, $url, $type;
		public function __construct($uid="-1", $label="", $type="", $helpText="", $placeHolderText="", $maxChars="40"){
			parent::__construct($uid, $label, $helpText, $placeHolderText);
			$this->type=$type;
			$this->maxChars=$maxChars;
			$this->iconStyle="fa-lg";
			$this->class = "btn btn-default";
			$this->style = "";
			$this->type = "button";
		}
		public function setURL($url){
			$this->url = $url;
		}
		public function setType($url){
			$this->type = $url;
		}
		public function InnerShow(){
			if($this->url!=""){
			$folder = explode("/", $_SERVER["REQUEST_URI"]);
			$_SystemRootRedirectName = $folder[1];	
			?>
			<a href="../<?php echo $_SystemRootRedirectName."/".$this->url ?>"> 
			<?php } ?>
			<button type="<?php echo $this->type ?>" <?php echo $this->atributesString ?> input-type="<?php echo $this->type?>" input-id="<?php echo $this->uid?>" id="<?php echo $this->uid?>" name="<?php echo $this->name?>" class="<?php echo $this->class ?>"  style="<?php echo $this->style ?>" <?php if(!$this->active) echo "disabled" ?>>
				<?php
				if($this->innerElement==null){ ?>
					<?php if($this->icon=="") echo $this->label; else{ ?>
						<i class="fa <?php echo $this->icon." ".$this->iconStyle ?> "></i>
						<label style="<?php echo $this->labelStyle ?>" class="control-label button-label <?php echo $this->labelClass ?>"><?php echo $this->label;?></label>
					<?php } ?>
					<?php
				}else{
					if(gettype($this->innerElement)!="string"){
						if(get_parent_class($this->innerElement)=="formElement"){
							$this->innerElement->InnerShow();
						}
					}else{
						echo $this->innerElement;
					}
				}
				?>
			</button>
		<?php
			if($this->url!=""){
				echo "</a>";
			}
		}
		public function setIcon($icon, $style){
			$this->icon=$icon;
			$this->iconStyle=$style;
		}
		public function setLink($url){
			$this->url=$url;
		}
		function showLabel(){
			if($this->label!="" && $this->showLabel){
				if($this->icon==""){
				?>
				<label style="<?php echo $this->labelStyle ?>" class="control-label <?php echo $this->labelClass ?>"><?php //echo $this->label;?></label>
				<?php
				}
			}
		}
	}

	class RowElement{
		public $class, $element;
		public function __construct($rowClass, $element){
			$this->class = $rowClass;
			$this->element = $element;
		}
		public function setActive($status){
			if(is_array($this->element)){
				foreach($this->element as $element){
					if(method_exists($element, "setActive")){
						$element->setActive($status);
					}
				}
			}else{
				if(is_object($this->element) && method_exists($this->element, "setActive")){	
					$this->element->setActive($status);
				}
			}
		}
	}
	
	class DivElement extends formElement{
		public $class;
		public $element, $showPanel, $innerHTML, $panelStyle, $panelClass;
		public $rows, $LeftColumnClass, $RightColumnClass;
		public $itemListLeft, $itemListMiddle, $itemListMiddleBefore, $itemListRight;
		public function __construct(){
			$this->class = "panel panel-default panel-body";
			$this->itemListLeft=array();
			$this->itemListRight=array();
			$this->itemListMiddle=array();
			$this->itemListMiddleBefore=array();
			$this->atributes=array();
			$this->panelStyle="";
			$this->panelClass="";
			$this->showPanel=false;
			$this->visible=true;
			$this->title="";
			$this->innerHTML="";
			$this->removePercentage=false;
			$this->rows=array();
			$this->percentage = " width: 100%; ";
			$this->LeftColumnClass = "";
			$this->RightColumnClass = "6";
		}
		public function showPanel($bool){
			$this->showPanel=$bool;
		}
		public function setClass($value){
			$this->class=$value;
		}
		public function setRightColumn($value){
			if(is_array($value)){
				$this->itemListRight=$value;
			}else{
				$this->itemListRight = array($value);
			}
		}
		public function setLeftColumn($value){
			if(is_array($value)){
				$this->itemListLeft=$value;
			}else{
				$this->itemListLeft = array($value);
			}
		}
		public function setMiddleColumn($value){
			if(is_array($value)){
				$this->itemListMiddle=$value;
			}else{
				$this->itemListMiddle = array($value);
			}
		}
		public function setMiddleBeforeColumn($value){
			if(is_array($value)){
				$this->itemListMiddleBefore=$value;
			}else{
				$this->itemListMiddleBefore = array($value);
			}
		}
		public function setColumns($value){
			$this->itemListLeft = array();
			$this->itemListRight = array();
			$bool=false;
			foreach($value as $atual){
				if(!$bool){
					array_push($this->itemListLeft, $atual);
				}else{
					array_push($this->itemListRight, $atual);
					}
				
				$bool=!$bool;
			}
		}
		
		public function setTitle($name){
			$this->title=$name;
			if(trim($name)!=""){
				$this->showPanel(true);
			}
		}
		public function setPanelStyle($name){
			$this->panelStyle = $name;
		}
		public function setPanelClass($name){
			$this->panelClass = $name;
		}
		public function doRemovePercentage(){
			$this->removePercentage=true;
		}
		public function Show($toString=false){
			if($toString)
                $this->InnerShowToString();
			else
				$this->InnerShow();
		}
		public function InsertRow($classRow, $elements){
			array_push($this->rows, new RowElement($classRow, $elements));
		}
		function setActive($status){
			$total = array_merge($this->itemListLeft, $this->itemListMiddle, $this->itemListMiddleBefore, $this->itemListRight, $this->rows);
			foreach($total as $item){
				if(get_parent_class($item)=="formElement"){
					$item->setActive($status);
				}else{
					if(get_class($item)=="RowElement"){
						$item->setActive($status);
					}
				}
			}
		}
		public function setPercentage($val = "100%"){
			$this->percentage = " width: ".$val."; ";
		}
		public function InnerShow(){
			if($this->showPanel==true){
			?>
			<div class="panel panel-default <?php echo $this->panelClass ?>" style="<?php if($this->panelStyle!="") echo $this->panelStyle; if($this->visible==false) echo "display: none" ?>">
				<?php if($this->title!=""){ ?>
				<div class="panel-heading"><?php echo $this->title ?></div>
				<?php } ?>
				<div class="panel-body">
					<?php
					}
					//<!--<div id="<?php echo $this->uid? >" <?php echo $this->atributesString ? > class="<?php echo $this->class ? >" style="<?php if($this->style!="") echo $this->style; if(!$this->removePercentage){ ? > width:100%; <?php } if($this->visible==false && $this->showPanel==false) echo "display: none" ? >
					?>					
						<div id="<?php echo $this->uid?>" <?php echo $this->atributesString ?> class="<?php echo $this->class ?>" style="<?php if($this->style!="") echo $this->style; if(!$this->removePercentage){ echo $this->percentage; } if($this->visible==false && $this->showPanel==false) echo "display: none" ?>">
						<?php 
						if(count($this->itemListMiddleBefore)>0){ ?>
							<div class="col-lg-12">
								<?php
								foreach($this->itemListMiddleBefore as $atual){							
								
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						}
						/* if(count($this->itemListLeft)>0){?>
							<div class="col-lg-<?php if(count($this->itemListRight)>0){ echo "6"; }else{ echo "12"; }?>">
								<?php
								foreach($this->itemListLeft as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
							<?php 
						}
						if(count($this->itemListRight)>0){?>
							<div class="col-lg-6">
								<?php
								foreach($this->itemListRight as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						} */
						if(count($this->itemListLeft)>0){/*<div class="col-lg-<?php if(count($this->itemListRight)>0){ echo "6"; }else{ echo "12"; }*/?>
							<div class="col-lg-<?php if(count($this->itemListRight)>0){ if($this->LeftColumnClass != ""){ echo $this->LeftColumnClass;}else{echo "6";} }else{ if($this->LeftColumnClass != ""){echo $this->LeftColumnClass;}else{echo "12";}}?>">							
								<?php
									foreach($this->itemListLeft as $atual){
										if(method_exists($atual, "Show")){
											$atual->Show();
										}else if(method_exists($atual, "InnerShow")){
											$atual->InnerShow();
										}else{
											echo $atual;
										}
									}
								?>
							</div>
							<?php 
						}
						if(count($this->itemListRight)>0){//<div class="col-lg-6">?>							
							<div class="col-lg-<?php echo $this->RightColumnClass;?>">
								<?php
								foreach($this->itemListRight as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						}
						if(count($this->itemListMiddle)>0){?>
							<div class="col-lg-12">
								<?php
								foreach($this->itemListMiddle as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						}
						if(count($this->rows)>0){
							?><div class = "row"><?php
							foreach($this->rows as $row){ ?>
									<div class="col-lg-<?php echo $row->class?>">
										<?php
										if(is_array($row->element)){
											foreach($row->element as $rowElement){
												if(method_exists($rowElement, "Show")){
													$rowElement->Show();
												}else if(method_exists($rowElement, "InnerShow")){
													$rowElement->InnerShow();
												}else{
													echo $rowElement;
												}
											}
										}else{
											if(method_exists($row->element, "Show")){
												$row->element->Show();
											}else if(method_exists($row->element, "InnerShow")){
												$row->element->InnerShow();
											}else{
												echo $row->element;
											}
										}
										?>
									</div>
	
							<?php
							}
							?></div><?php
						}?>
						<?php if($this->innerHTML!=""){
							echo $this->innerHTML;
						}?>
					</div>
					<?php
					if($this->showPanel==true){
					?>
				</div>
			</div>
			<?php
			}
		}
		
		public function InnerShowToString(){
			if($this->showPanel==true){
			?>
			<div class="panel panel-default <?php echo $this->panelClass ?>" style="<?php if($this->panelStyle!="") echo $this->panelStyle; if($this->visible==false) echo "display: none" ?>">
				<?php if($this->title!=""){ ?>
				<div class="panel-heading"><?php echo $this->title ?></div>
				<?php } ?>
				<div class="panel-body">
					<?php
					}
					//<!--<div id="<?php echo $this->uid? >" <?php echo $this->atributesString ? > class="<?php echo $this->class ? >" style="<?php if($this->style!="") echo $this->style; if(!$this->removePercentage){ ? > width:100%; <?php } if($this->visible==false && $this->showPanel==false) echo "display: none" ? >
					?>					
						<div id="<?php echo $this->uid?>" <?php echo $this->atributesString ?> class="<?php echo $this->class ?>" style="<?php if($this->style!="") echo $this->style; if(!$this->removePercentage){ echo $this->percentage; } if($this->visible==false && $this->showPanel==false) echo "display: none" ?>">
						<?php 
						if(count($this->itemListMiddleBefore)>0){ ?>
							<div class="col-lg-12">
								<?php
								foreach($this->itemListMiddleBefore as $atual){							
								
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						}
						/* if(count($this->itemListLeft)>0){?>
							<div class="col-lg-<?php if(count($this->itemListRight)>0){ echo "6"; }else{ echo "12"; }?>">
								<?php
								foreach($this->itemListLeft as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
							<?php 
						}
						if(count($this->itemListRight)>0){?>
							<div class="col-lg-6">
								<?php
								foreach($this->itemListRight as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						} */
						if(count($this->itemListLeft)>0){/*<div class="col-lg-<?php if(count($this->itemListRight)>0){ echo "6"; }else{ echo "12"; }*/?>
							<div class="col-lg-<?php if(count($this->itemListRight)>0){ if($this->LeftColumnClass != ""){ echo $this->LeftColumnClass;}else{echo "6";} }else{ if($this->LeftColumnClass != ""){echo $this->LeftColumnClass;}else{echo "12";}}?>">							
								<?php
									foreach($this->itemListLeft as $atual){
										if(method_exists($atual, "Show")){
											$atual->Show();
										}else if(method_exists($atual, "InnerShow")){
											$atual->InnerShow();
										}else{
											echo $atual;
										}
									}
								?>
							</div>
							<?php 
						}
						if(count($this->itemListRight)>0){//<div class="col-lg-6">?>							
							<div class="col-lg-<?php echo $this->RightColumnClass;?>">
								<?php
								foreach($this->itemListRight as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						}
						if(count($this->itemListMiddle)>0){?>
							<div class="col-lg-12">
								<?php
								foreach($this->itemListMiddle as $atual){
									if(method_exists($atual, "Show")){
										$atual->Show();
									}else if(method_exists($atual, "InnerShow")){
										$atual->InnerShow();
									}else{
										echo $atual;
									}
								}
								?>
							</div>
						<?php 
						}
						if(count($this->rows)>0){
							?><div class = "row"><?php
							foreach($this->rows as $row){ ?>
									<div class="col-lg-<?php echo $row->class?>">
										<?php
										if(is_array($row->element)){
											foreach($row->element as $rowElement){
												if(method_exists($rowElement, "Show")){
													$rowElement->Show();
												}else if(method_exists($rowElement, "InnerShow")){
													$rowElement->InnerShow();
												}else{
													echo $rowElement;
												}
											}
										}else{
											if(method_exists($row->element, "Show")){
												$row->element->Show();
											}else if(method_exists($row->element, "InnerShow")){
												$row->element->InnerShow();
											}else{
												echo $row->element;
											}
										}
										?>
									</div>
	
							<?php
							}
							?></div><?php
						}?>
						<?php if($this->innerHTML!=""){
							echo $this->innerHTML;
						}?>
					</div>
					<?php
					if($this->showPanel==true){
					?>
				</div>
			</div>
			<?php
			}
		}
	}
	
	class HeadingElement extends formElement{
		public $number;
		public function __construct($uid="", $number="2", $value="Header"){
			$this->uid = $uid;
			$this->value=$value;
			$this->number=$number;
			$this->visible=true;
		}
		public function InnerShow(){
        $this->visible=true;
			?>
			<div id="<?php echo $this->uid?>" name="<?php echo $this->uid?>">
				<h<?php echo $this->number?>>
				   <?php echo $this->value?>
				</h<?php echo $this->number?>>
			</div>
			<?php
		}
	}

	class HeaderElement extends formElement{
		public $col, $H;
		public function __construct($label="", $H=2){
			$this->col = "12";
			$this->value = $label;
			$this->H = $H;
			$this->showLabel = false;
		}
		
		public function SetColNumber($value){
			$this->col = $value;
		}
		public function SetHSize($value){
			$this->H = $value;
		}
		public function InnerShow(){
			?>		
			<?php if($this->showContainerDiv){ ?><div asdasdsdsa="Asdasdasd" class="col-lg-<?php echo $this->col; ?>" id="<?php echo $this->uid?>" name="<?php echo $this->name?>" style="<?php echo $this->style?>"><?php } ?>
				<h<?php echo $this->H ?>  class="page-headers <?php echo $this->class ?>" <?php if(!$this->showContainerDiv){ ?>style="<?php echo $this->style ?>" id="<?php echo $this->uid?>" <?php } ?>		><?php echo $this->value ?></h<?php echo $this->H ?>>
			<?php if($this->showContainerDiv){ ?></div>	<?php } ?>		
			<?php
		}
		public static function HeaderDefault($label="", $size=2, $uid=-1, $print=true){
			//$div = new SimpleTagElement();
			/*$div->insertRow("12", array(*/$div = new HeaderElement($label, $size);
			//$div->setLabel(null);
			$div->setId($uid);
			$div->showContainerDiv(false);
			$div->setClass("row col-lg-12");
			if($print){
				$div->Show();
			}else{
				return $div;
			}
		}
	}

	class InputElement extends formElement{
		public $maxChars, $autoCompleteElements = array();
		public function __construct($uid="-1", $label="", $type="", $helpText="", $placeHolderText="", $maxChars="40"){
			parent::__construct($uid, $label, $helpText, $placeHolderText);
			$this->type=$type;
			$this->maxChars=$maxChars;
			$this->name = $this->uid;
		}
		public function setAutoComplete($values){
			$this->autoCompleteElements=$values;
		}
		public function InnerShow(){
		?>
			<input input-type="<?php echo $this->type?>" id="<?php echo $this->uid?>" <?php echo $this->atributesString?> maxlength="<?php echo $this->maxChars?>" name="<?php echo $this->name?>" class="form-control <?php echo $this->type?> <?php echo $this->class?>" placeholder="<?php echo $this->placeHolderText?>" value ="<?php echo $this->value?>" <?php if(!$this->active) echo 'disabled'?> <?php if($this->required==true) echo ' required="required"'?> style=" <?php if($this->visible==false) echo 'display: none;'?> <?php if($this->style!="") echo $this->style; ?>"/>	
		<?php
			if(count($this->autoCompleteElements)>0){
				?>
				<script>
					$(function() {
						var availableTags = [
							<?php
								$string = "";
								foreach($this->autoCompleteElements as $key=>$element){
									$string .= "\"".$key."\",";
								}
								rtrim($string, ",");
								echo $string;
							?>
						];
						$("#<?php echo $this->uid?>").autocomplete({ source: availableTags });
					});
				</script>
				<?php
			}
		}
	}	

	class SelectElement extends formElement{
		public $labels;
		public function __construct($uid="", $label="", $helpText="", $placeHolderText="", $labels=array(0=>"Radio1"), $type=0){
			parent::__construct($uid, $label, $helpText, $placeHolderText);
			$this->type=$type;
			$this->labels=array();
			$this->class="form-control";
			$this->style="";

			foreach($labels as $key=> $atual){
				$this->labels[$key]=$atual;
			}
		}

		public function InnerShow(){
			//Retirado 'required' do select e colocado no input respectivo
			?>
			<select <?php if($this->type==1) echo "multiple"?> class="<?php echo $this->class ?>" <?php echo $this->atributesString ?> style="<?php echo $this->style ?>" <?php if(!$this->active) echo 'disabled'?> id="<?php echo $this->uid?>_select" >
			<?php if($this->type!=1) echo '<option value="" disabled selected> '.$this->placeHolderText.'</option>'?>
			<?php
			$i=0;
			foreach($this->labels as $key=> $atual){  // 0,  1
				?><option value="<?php echo $atual?>" <?php if($atual==$this->value) echo "selected";?> ><?php
				echo $key;
				?></option><?php
				$i++;
			}
			?>
			</select>
			<input type="hidden" id="<?php echo $this->uid?>" name="<?php echo $this->uid?>" value="<?php echo $this->value?>" <?php if($this->required) echo ' required="required"'?>>
			<?php
		}
	}
	
	class Cards extends formElement{
		public $TextMain,$ClassTextSub,$colorClardbyClass, $TextSub,$idButton, $disable,$elementAddButton ,$TextButton,$colorButtonbyClass, $TextLeftButton, $TextLeftButtonExist = false, $imagemBase = null, $imageExiste;
		public function __construct($uid="-1", $label=" ", $helpText="", $placeHolderText=""){
			parent::__construct($uid, $label, $helpText, $placeHolderText);
			$this->disable = "";
			$this->classSpaceFoolter = "justify-content-between";
		}
		public function setTextMain($value){
			$this->TextMain = $value;
		}
		public function setTextSub($value){
			$this->TextSub = $value;
		}
		public function setClassTextSub($value){
			$this->ClassTextSub = $value;
		}
		public function setIdButton($value){
			$this->idButton = $value;
		}
		public function setElementAddButton($value){
			$this->elementAddButton .= " ".$value;
		}
		public function setTextButton($value){
			$this->TextButton = $value;
		}
		public function setColorButtonbyClass($value){
			$this->colorButtonbyClass = $value;
		}
		public function setColorCardbyClass($value){
			$this->colorClardbyClass = $value;
		}
		public function setTextLeftButtonExist($value = false){
			$this->TextLeftButtonExist = $value;
		}
		public function setTextLeftButton($value){
			$this->TextLeftButton = $value;
		}
		public function setImageExiste($value){
			$this->imageExiste = $value;
		}

		public function setDisable(){
			$this->disable = "disabled";
		}
		public function setimagemBase($value,$nameFile){
			$base64 = $value;
			$decoded = base64_decode($base64);
			$file = dirname(__DIR__)."/images/modalidades/".$nameFile;
			file_put_contents($file, $decoded);
			
			if (file_exists($file)) {
				$this->imagemBase = $nameFile;
			}
		}
		public function InnerShow(){
			?>
			<div class="card <?php echo $this->colorClardbyClass?>" style="width: 18rem;">
				<?php if($this->imageExiste == 1) echo '<img class="card-img-top" src="/images/modalidades/'.$this->imagemBase.'" alt="Card image">'?>
				<div class="card-body" id="<?php echo $this->uid?>">
					<h5 class="card-title"><?php echo $this->TextMain?></h5>
					<p class="card-text <?php echo $this->ClassTextSub?>"><?php echo $this->TextSub?></p>
					<div class="d-flex <?php echo $this->classSpaceFoolter?>">
						<a href="#" input-send="buttonCard" input-id="<?php echo $this->idButton?>"  <?php echo $this->elementAddButton?> class="btn <?php echo $this->colorButtonbyClass.' '.$this->disable?>"><?php echo $this->TextButton?></a>
						<span href="" class="btn  <?php echo $this->colorButtonbyClass.' '.$this->disable?>"><?php if($this->TextLeftButtonExist) echo $this->TextLeftButton?></span>
					</div>
				</div>
				</div>
			<?php
		}
	}

	class HTMLElement extends formElement{
		public $html;
		public function __construct($html){
			$this->html = $html;
		}
		public function setHtml($value){
			$this->html = $value;
		}
		public function InnerShow(){
			echo $this->html;
		}
	}
	
	class GridBlock extends formElement{
		public $colspan, $element,  $isHeader, $color;
		public function __construct($element, $isHeader=false, $colspan=1){
			$this->element = $element;
			$this->colspan = $colspan;
			$this->isHeader = $isHeader;
		}
		public function InnerShow(){
			$this->Show();
		}
		public function Show(){
			$tag = ($this->isHeader)?"th":"td";
			?>
		<<?php echo $tag?> colspan="<?php echo $this->colspan?>" id="<?php echo $this->uid?>" <?php echo $this->atributesString ?> class="<?php echo $this->class ?>" style="display:table-cell; <?php if($this->color!=""){?> background-color:<?php echo $this->color; }?>; <?php if($this->style!="") echo $this->style; ?>">
			<?php
			if(is_array($this->element)){
				foreach($this->element as $ele){
					if(get_parent_class($ele)=="formElement"){
						$ele->InnerShow();
					}else{
						if(trim($ele)=="") $ele="-";
						echo $ele;
					}
				}
			}else{
				if(get_parent_class($this->element)=="formElement"){
					$this->element->InnerShow();
				}else{
					if(is_object($this->element)){
						if(get_class($this->element)=="Form"){
							$this->element->Show();
						}
					}else{
						if(trim($this->element)=="") $this->element="-";
						echo $this->element;
					}
				}
			}
			?>
			</<?php echo $tag ?>>
			<?php
		}
	}

	class GridLine extends formElement{
		public $color, $columns;
		public $cellColor = array();
		public function __construct($columns=array()){
			$color = "";
			$go = true;
			foreach($columns as $column){
				if(!is_object($column) || is_object($column) && get_class($column)!="GridBlock"){
					$go=false;
				}
			}
			$this->columns = ($go)?$columns:array();
			if(!$go) throw new Exception('Erro. Só são permitidos objetos do tipo "GridBlock"');
		}
		public function setColumns($column){
			$this->columns = $column;
		}
		public function setColor($color){
			$this->color = $color;
		}
		public function InnerShow(){
			$this->Show();
		}
		public function setCellColor($color=array()){
			$this->cellColor = $color;
		}
		public function Show(){
			$counter = 0;
			?><tr 
				<?php echo $this->atributesString ?> 
				class="<?php echo $this->class ?>" 
				style="<?php echo $this->style ?>"><?php
				foreach($this->columns as $column){	
					$column->color = $this->color;			
					if(count($this->cellColor) > 0 && is_array($this->cellColor)){
						if(array_key_exists($counter, $this->cellColor)){
							if($this->cellColor != ""){
								$column->color = $this->cellColor[$counter];
							}
						}
					}
					$counter += 1;				
					$column->Show();
				}
			echo "</tr>";
		}
		/* public function Show(){
			?><tr <?php echo $this->atributesString ?> class="<?php echo $this->class ?>" style="<?php echo $this->style ?>"><?php
			foreach($this->columns as $column){
				$column->color = $this->color;
				$column->Show();
			}
			echo "</tr>";
		} */
		public function setHeader($bool=true){
			foreach($this->columns as $col){
				$col->isHeader = $bool;
			}
		}
		public static function NewFromString(){
			$params = func_get_args();
			$columns = array();
			foreach($params as $column){
				array_push($columns, new GridBlock($column));
			}
			return new GridLine($columns);
		}
		public static function NewFromArray($array, $bold = false){
			$columns = array();
			foreach($array as $column){
				if($bold)
					$column = "<b>".$column."</b>";
				array_push($columns, new GridBlock($column));
			}
			return new GridLine($columns);
		}
	}

	class GridElement extends formElement{
		public $matrix, $noDataLabel, $deletavel;
		public function __construct($matrix, $noDataLabel="No data to show"){
			foreach($matrix as $atual){
				if(is_array($atual)){
					foreach($atual as $array1){
						if(!is_array($atual)){
							echo("There's an error on creating the grid");
						}
					}
				}
			}
			$this->matrix=$matrix;
			$this->noDataLabel=$noDataLabel;
			$this->deletavel=false;
		}
		public function setDeletable($value){
			$this->deletavel=(bool)$value;
		}
		public function setAjax($path){
			$this->addAttribute(array("ajaxPath"=>$path));
			$this->addAttribute("ajax");
		}
		public function InnerShow(){
			?>
			<table id="<?php echo $this->uid?>" <?php echo $this->atributesString ?> class="table table-striped table-bordered table-hover <?php echo $this->class ?>" style="<?php if($this->style!="") echo $this->style; ?>" >
			<?php
			$count=0;
			foreach($this->matrix as $atual){
				if($count==0){
					?>
						<thead>
					<?php }else if($count==1){
					?>
						<tbody>
					<?php
				}
				$atual->Show();
				if($count==0){
					?>
							</thead>
					<?php
				}
				$count++;
			}
			?>
				</tbody>
			</table>
			<?php
		}
		public static function generatePrototype($nomesColunas=array(),$ajaxPath,$idTable='dataTables',$returnTable = true,$centerCell = true){
            $arrayBlock = array();
            $lines = array();
			foreach($nomesColunas as $header){
				$block = new GridBlock($header); 
				if($centerCell)
					$block->setStyle("text-align:center;vertical-align: middle;");
				array_push($arrayBlock,$block);
			}
            $line = new GridLine($arrayBlock);
            $line->setHeader(true);
			if(!$returnTable)
				return $line;
            array_push($lines,$line);
            $grid = new GridElement($lines);
            $grid->setID($idTable);
            $grid->setAjax($ajaxPath);
            return $grid;
		}
	}
	
	class SeparatorElement extends InputElement{
		public $number;
		public function __construct($number=1){
			$this->number=$number;
			$this->visible=true;
		}
		public function InnerShow(){
			for($i=0; $i<$this->number; $i++){
				?>
				<div class="col-lg-12" temp="temp" style="/* visibility:hidden; */height: 1px;margin: 9px 0;overflow: hidden;background-color: #e5e5e5;"></div>
				<?php
			}
		}
	}

?>


	