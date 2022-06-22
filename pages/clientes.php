<?php 
	require_once("../config/config.php");
    if(isset($_SESSION))
        $host = $_SESSION['HTTP_HOST'];
    else
        $host = '';
	if(!isset($tipoConfigPagina)){
		$tipoConfigPagina = Utility::getCurrentPage();
        $dataReceived = (file_get_contents("php://input"));
        if($dataReceived!="");
            $session = json_decode($dataReceived);

	}
    switch($tipoConfigPagina){
        case "ajax":
			if(session_id() == "")
				session_start();				
			$case = Utility::getVariable("case", INPUT_GET);
			switch($case){
                case 1: // Funcionario Adm
                    $draw = Utility::getVariable("draw");
                    $start = Utility::getVariable("start");
                    $length = Utility::getVariable("length");
                    $order = Utility::getVariable("order");
                    $busca = Utility::getVariable("search");
					$busca = explode(" ", $busca["value"]);
                    $search = array("busca"=>$busca,"start"=>$start,"length"=>$length,"order" => $order);

                    $columns = array('i.nomeCompleto','ps.nomePlano','cp.dataInicio','cp.dataVencimento','pag.datePG');
                    $retorno = array();
                    $temp = array();
                    $querySql = "   SELECT cp.id as clienteID, i.nomeCompleto,cp.dataInicio
                                    FROM ".Cliente::$database." cp 
                                    JOIN ".Usuarioo::$database." u ON u.id = cp.idUsuario
                                    JOIN ".InfoDadoss::$database." i ON i.id = u.idInfoDados
                                    WHERE cp.ativo = 1";
                    Connection::query($querySql);
                    $numeroRegistros = Connection::$numRows;
                    $orderSend = array();
                    if(count($order) > 0){
                        $columnsOrder = array('i.nomeCompleto','cp.dataInicio','i.nomeCompleto',"","");
                        foreach($order as $orderUni){
                            $orderAux = new stdClass();
                            $keyAux =  $columnsOrder[$orderUni['column']];
                            $orderAux->$keyAux = $orderUni['dir'];
                            $orderSend = ($orderAux);
                        }
                    }
                    $querySql = Utility::formatQueryToSearchForColumns($querySql,$columns,$busca,$start,$length,$orderSend);
                    $fetch = Connection::query($querySql);
                    $retorno = array("totalValues"=>$numeroRegistros,"filteredValues"=>array());
                    if(Connection::$numRows>0){
                        foreach($fetch as $rs){
                            foreach ($rs as $key=>$value){
                                $temp[$key]=$value;
                            }
                            array_push($retorno['filteredValues'],$temp);
                        }
                    }
                    $data = array("data"=>array(), "draw"=>$draw,"iTotalDisplayRecords"=>$retorno['totalValues'], "iTotalRecords"=>$retorno['totalValues'], "success"=>true);
                    $retorno = $retorno['filteredValues'];
                    foreach ($retorno as $tableLine){
                        $status = 'Ativo';
                        $buttons = ' <button type="button" id="infoPerito" title="Informações de Contato" class="btn btn-inverse-primary btn-rounded btn-icon " data-idCliente="'.$tableLine['clienteID'].'" >
                                    <i class="mdi mdi-magnify-plus-outline" trigger = "detalhes-processo"></i>
                                  </button>
                                  <button type="button" name="relatorioPeritoAnalise" title="Relatório Clientes" class="btn btn-inverse-primary btn-rounded btn-icon" data-idCliente="'.$tableLine['clienteID'].'" >
                                    <i class="mdi mdi-table-large" trigger = "detalhes-processo"></i>
                                  </button>';
                        array_push($data['data'],array(
                                                $tableLine['nomeCompleto'],
                                                Utility::dateFormatToBR($tableLine['dataInicio']),
                                                $status,
                                                $buttons));
                    }
                    echo json_encode($data);
                    exit();
                break;
                case 2: // Funcionario Personal
                    $draw = Utility::getVariable("draw");
                    $start = Utility::getVariable("start");
                    $length = Utility::getVariable("length");
                    $order = Utility::getVariable("order");
                    $busca = Utility::getVariable("search");
					$busca = explode(" ", $busca["value"]);
                    $search = array("busca"=>$busca,"start"=>$start,"length"=>$length,"order" => $order);

                    $columns = array('i.nomeCompleto','ps.nomePlano','cp.dataInicio','cp.dataVencimento','pag.datePG');
                    $retorno = array();
                    $temp = array();
                    $querySql = "   SELECT cp.id as clienteID, i.nomeCompleto,cp.dataInicio
                                    FROM ".Cliente::$database." cp 
                                    JOIN ".Usuarioo::$database." u ON u.id = cp.idUsuario
                                    JOIN ".InfoDadoss::$database." i ON i.id = u.idInfoDados
                                    WHERE cp.ativo = 1";
                    Connection::query($querySql);
                    $numeroRegistros = Connection::$numRows;
                    $orderSend = array();
                    if(count($order) > 0){
                        $columnsOrder = array('i.nomeCompleto','cp.dataInicio','i.nomeCompleto',"","");
                        foreach($order as $orderUni){
                            $orderAux = new stdClass();
                            $keyAux =  $columnsOrder[$orderUni['column']];
                            $orderAux->$keyAux = $orderUni['dir'];
                            $orderSend = ($orderAux);
                        }
                    }
                    $querySql = Utility::formatQueryToSearchForColumns($querySql,$columns,$busca,$start,$length,$orderSend);
                    $fetch = Connection::query($querySql);
                    $retorno = array("totalValues"=>$numeroRegistros,"filteredValues"=>array());
                    if(Connection::$numRows>0){
                        foreach($fetch as $rs){
                            foreach ($rs as $key=>$value){
                                $temp[$key]=$value;
                            }
                            array_push($retorno['filteredValues'],$temp);
                        }
                    }
                    $data = array("data"=>array(), "draw"=>$draw,"iTotalDisplayRecords"=>$retorno['totalValues'], "iTotalRecords"=>$retorno['totalValues'], "success"=>true);
                    $retorno = $retorno['filteredValues'];
                    foreach ($retorno as $tableLine){
                        $status = 'Ativo';
                        $buttons = ' <button type="button" id="infoPerito" title="Informações de Contato" class="btn btn-inverse-primary btn-rounded btn-icon " data-idCliente="'.$tableLine['clienteID'].'" >
                                    <i class="mdi mdi-magnify-plus-outline" trigger = "detalhes-processo"></i>
                                  </button>';
                        array_push($data['data'],array(
                                                $tableLine['nomeCompleto'],
                                                Utility::dateFormatToBR($tableLine['dataInicio']),
                                                $status,
                                                $buttons));
                    }
                    echo json_encode($data);
                    exit();
                break;                           
                case 3: // Info Cliente

                    $idCliente = Utility::getVariable("idCliente");

                    $infoCliente = Cliente::getInfoDados($idCliente);

                    $response = array();
                    $response["nome"] = explode(" " ,$infoCliente["nomeCompleto"])[0]; 
                    $response["status"] = 200;
                    if($infoCliente["telefone"] != null)
                        $response["telefone"] = $infoCliente["telefone"]; 
                    if($infoCliente["telefone"] != null)
                        $response["email"] = $infoCliente["email"]; 

                    echo json_encode($response);
                    exit();
                break;

			}
		break;

        default:
            $json = json_encode(array("id"=>$session->id));
            $user = new Usuarioo($session->id);
            $tipoAcesso = $user->idTipoAcesso_id;

            switch ($tipoAcesso){
                case 1:
                    $arrayBlock = array();
                    $lines = array();
                    $block = new GridBlock("Cliente");
                    array_push($arrayBlock,$block);
                    $block = new GridBlock("Inicio"); 
                    array_push($arrayBlock,$block);
                    $block = new GridBlock("Status"); 
                    array_push($arrayBlock,$block);
                    $block = new GridBlock("-"); 
                    array_push($arrayBlock,$block);

                    $line = new GridLine($arrayBlock);
                    array_push($lines,$line);
                    $div = new DivElement();
                    $grid = new GridElement($lines);
                    $grid->setID("myTable");
                    $grid->setAjax($host."/pages/clientes.php?status=ajax&case=1");
                    $div->InsertRow('12',$grid);
                    $div->show();
                    ?>
                    <div class="modal fade" id="modal" role="dialog" style="height: 100vh">
                        <div class="modal-dialog modal-lg" style = "width:100%;max-width:90vw">
                            <div class="modal-content">
                                <div class="modal-header" style="position:relative; top: 0; left: 0; right: 17px; z-index: 10; background: white;">
                                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                                    <h4 class="modal-title"></h4>
                                </div>
                                <div style="overflow-y: scroll; height: 80%;" class="modal-body">						
                                                                                                
                                </div>
                                <div class="modal-footer" style="position:relative; bottom: 0; left: 0; right: 17px; z-index: 10; background: white;">
                                    <button type="button" class="btn btn-success" data-dismiss="salvaForm">Salvar</button>
                                    <button type="button" class="btn btn-default" data-bs-dismiss="modal">Sair</button>
                                </div>
                            </div>
                        </div>
                    </div>
                    <?php
                break;

                case 3:
                    $arrayBlock = array();
                    $lines = array();
                    $block = new GridBlock("Cliente");
                    array_push($arrayBlock,$block);
                    $block = new GridBlock("Inicio"); 
                    array_push($arrayBlock,$block);
                    $block = new GridBlock("Status"); 
                    array_push($arrayBlock,$block);
                    $block = new GridBlock("-"); 
                    array_push($arrayBlock,$block);

                    $line = new GridLine($arrayBlock);
                    array_push($lines,$line);
                    $div = new DivElement();
                    $grid = new GridElement($lines);
                    $grid->setID("myTable");
                    $grid->setAjax($host."/pages/clientes.php?status=ajax&case=1");
                    $div->InsertRow('12',$grid);
                    $div->show();
                break;
            }

            Utility::InsertJavascriptFile("clientes");
        break;
    }
?>

