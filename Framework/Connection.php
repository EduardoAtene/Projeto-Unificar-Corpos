<?php
    abstract class Connection{
		private static $usuario;
		private static $senha;
		private static $host;
		public static $liConn;
		public static $iniFile;
		public static $conditions, $database, $table, $order, $group, $joins, $lastInsertedId, $tableType, $debug, $numFields, $numRows;
		public static function initialize($dataBase="unificar_corpo"){	
			Connection::$iniFile="/configs/ConnectionConfig.ini";
			if(!file_exists(dirname($_SERVER['DOCUMENT_ROOT']).Connection::$iniFile)){
				die("Error. Configuration file not found: <br> ".dirname($_SERVER['DOCUMENT_ROOT']).Connection::$iniFile);
			}
			Connection::$usuario=Connection::getUsuario();
			Connection::$senha=Connection::getSenha();
			Connection::$database=$dataBase;
			Connection::$host=Connection::getHost();
			// var_dump(Connection::$usuario,Connection::$senha,Connection::$host,Connection::$database);
			// exit();
			try{
				Connection::$liConn=new mysqli(Connection::$host, Connection::$usuario, Connection::$senha, Connection::$database) or die("Error. Configuration file with incorrect values.<br>Could not connect to database.");
				Connection::$liConn->set_charset("utf8");
				Connection::$liConn->query("SET NAMES 'utf8'");
				Connection::$liConn->query("SET character_set_connection=utf8");
				Connection::$liConn->query("SET character_set_client=utf8");
				Connection::$liConn->query("SET character_set_results=utf8");
			}catch(Exception $e){
				echo $e->getMessage();
			}
			
		}
		
		public static function ReInitialize($dataBase="sistemaperito"){
			try{
				Connection::$database = $dataBase;
				Connection::$liConn->select_db($dataBase);
			}catch(Exception $e){
				echo $e->getMessage();
			}
		}

		public static function getServer(){
			$ini=parse_ini_file(dirname($_SERVER['DOCUMENT_ROOT']).Connection::$iniFile);
			return $ini['serverSQL'];
		}

		public static function getHost(){
			$ini=parse_ini_file(dirname($_SERVER['DOCUMENT_ROOT']).Connection::$iniFile);
			return $ini['hostSQL'];
		}

		public static function getUsuario(){
			$ini=parse_ini_file(dirname($_SERVER['DOCUMENT_ROOT']).Connection::$iniFile);
			return $ini['usuarioSQL'];
		}

		public static function getSenha(){
			$ini=parse_ini_file(dirname($_SERVER['DOCUMENT_ROOT']).Connection::$iniFile);
			return $ini['senhaSQL'];
		}

		static function setTableType($bool){
			Connection::$tableType=$bool;
		}

		static function setConditions($array){
			Connection::$conditions=str_replace("where", "", $array);
		}

		static function setJoins($array){
			Connection::$joins=$array;
		}

		static function setTable($array){
			Connection::$table=$array;
		}

		static function setOrder($string){
			Connection::$order=$string;
		}

		static function setGroup($string){
			Connection::$group=$string;
		}
		
		static function setDebug(){
			Connection::$debug=true;
		}

		static function query($query, $showError=true, $derp=0){
			if($derp==1){
				echo " - ", $query, " - ";
			}
			$obj=null;
			$obj=Connection::$liConn->query($query);
			if((Connection::$liConn->error!="" && $showError)){
				echo "Error: ".Connection::$liConn->error."<BR>Query: ".$query." <br>";
				return null;
			}
			Connection::$lastInsertedId=Connection::$liConn->insert_id;
			Connection::$numFields=Connection::$liConn->field_count;
			if(is_object($obj)){
				Connection::$numRows=$obj->num_rows;
			}
			return $obj;
		}

		static function consult(){
			$rows=array();
			foreach(func_get_args() as $n){
				if(is_string($n)){
					array_push($rows, $n);
				}
			}
			//$pagina=Utility::getVariable("pagina", INPUT_GET, FILTER_REQUIRE_SCALAR);
			$conditions=Connection::$conditions;
			$joins=Connection::$joins;
			$table=Connection::$table;
			$order=Connection::$order;
			$group=Connection::$group;
			$tableType=Connection::$tableType;
			Connection::$numRows=-1;
			Connection::$numFields=-1;
			// if($table==""){
			// 	$table=$pagina;
			// }
			if($tableType==true){
				$id="id_".$table.", ";
			}else{
				$id="";
			}
			$query="select ".((count($rows)>0) ? $id.implode(", ", $rows) : "*")." from ".$table." ".substr($table, 0, 1);
			$query .= " ".$joins." ";
			if(is_array($conditions)){
				if(count($conditions)>0){
					$query.=" where ".implode(' and ', array_map(function ($v, $k){ return $k.'=\''.$v.'\''; }, $conditions, array_keys($conditions)));
				}
			}else if($conditions!=""){
				$query.=" where ".$conditions;
			}
			if($group!=""){
				$query.=" group by ".$group;
			}
			if($order!=""){
				$query.=" order by ".$order;
			}
			if(Connection::$debug){
				echo $query;
			}
			$query=Connection::query($query);
			$queries=array();
			while($result=$query->fetch_array(MYSQLI_ASSOC)){
				array_push($queries, $result);
			}
			Connection::setConditions(array());
			Connection::setTable("");
			Connection::setOrder("");
			Connection::setGroup("");
			Connection::setJoins("");
			Connection::setTableType(true);
			Connection::$debug=false;
			return $queries;
		}
	
	}
