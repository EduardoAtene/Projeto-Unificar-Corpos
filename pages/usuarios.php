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
                case 1:
                    $response = $funcionarioArray = $clienteArray = $personalArray = array();
                    $funcionario = json_decode(Usuarioo::search(json_encode(array("idTipoAcesso_id"=>1)),true,"equal"));
                    if($funcionario->status =='200'){
                        $funcionarioArray = $funcionario->data;
                    }
                    $cliente = json_decode(Usuarioo::getAtivosId());
                    if($cliente->status =='200'){
                        $clienteArray = $cliente->data;
                    }
                    $personal = json_decode(Usuarioo::search(json_encode(array("idTipoAcesso_id"=>3)),true,"equal"));
                    if($personal->status =='200'){
                        $personalArray = $personal->data;
                    }
                    $usuariosArray = array_merge($funcionarioArray , $personalArray , $clienteArray);
                    foreach($usuariosArray as $usuarioId){
                        $user = new Usuarioo($usuarioId);
                        $info = new InfoDadoss($user->idInfoDados);
                        $acesso = new TipoAcessoo($user->idTipoAcesso_id);

                        $button = '<button type="button" name = "edit-user" class="btn btn-outline-secondary btn-rounded btn-icon">
                                        <i class="mdi mdi-pencil text-dark"></i>
                                    </button>';
                        array_push($response,array( $button,
                                                    $info->nomeCompleto,
                                                    $info->CPF,
                                                    $info->email,
                                                    $info->telefone,
                                                    $acesso->nome,
                                                    'Ativo'

                        ));
                    }
                    echo json_encode(array("data"=>$response));
                    exit();
                break;

                case 2:
                    ob_start(); // Start output buffering
                    if(!(Utility::getVariable("isedit"))){
                        $tipoAcesso = Utility::getVariable("tipoAcesso",INPUT_GET);
                        $response = array("header"=>"Cadastrar Novo Usuário ".ucfirst($tipoAcesso),"body"=>"","footer"=>"");
                        $usuario = new InputElement("usuario","Usuário","text","","usuário para fazer login","250");
                        $senha = new InputElement("senha","Senha","password","","senha","250");
                        $nome = new InputElement("nome","Nome Completo","text","","seu nome","250");
                        $cpf =new InputElement("cpf","CPF","number","","digite somente números");
                        $email = new InputElement("email","Email","email","","fulano@algo.com");
                        $tel = new InputElement("tel","Telefone de Contato Celular","number","","DDD + Número");
                        if($tipoAcesso != "cliente"){
                            $cargo =new InputElement("cargo","Cargo","text","","Digite o Cargo do Funcionário");
                        }
                        $divL = new DivElement();
                        $divL->InsertRow('6',$usuario);
                        $divL->InsertRow('6',$senha);
                        $divL->InsertRow('6',$nome);
                        $divL->InsertRow('6',$cpf);
                        $divL->InsertRow('6',$email);
                        $divL->InsertRow('6',$tel);
                        if($tipoAcesso != "cliente"){
                            $divL->InsertRow('6',$cargo);
                        }
                        $divL->show();
    
                        $response["status"] = 200;
                        $response["footer"] = '<button type="button" class="btn btn-success" data-dismiss="salvaForm" data-typeForm = "'.$tipoAcesso.'">Salvar</button> <button type="button" class="btn btn-default" data-bs-dismiss="modal">Sair</button>';
                        $response["body"] = ob_get_contents(); // Store buffer in variable
                    }
                    else{
                        $nome = Utility::getVariable("nomeCompleto");
                        $cpf = Utility::getVariable("cpf");

                        $search = json_decode(InfoDadoss::search(json_encode(array("nomeCompleto"=>$nome, "cpf"=>$cpf)),false));
                        try {
                            if($search->status==200){
                                $infoDados = new InfoDadoss(($search->data)[count($search->data)-1]);
                                $nome = new InputElement("nome","Nome Completo","text","","seu nome","250");
                                $nome->setValue($infoDados->nomeCompleto);
                                $nome->setActive(false);
                                $cpf =new InputElement("cpf","CPF","number","","digite somente números");
                                $cpf->setValue($infoDados->CPF);
                                $cpf->setActive(false);
                                $email = new InputElement("email","Email","email","","fulano@algo.com");
                                $email->setValue($infoDados->email);
                                $tel = new InputElement("tel","Telefone de Contato Celular","number","","DDD + Número");
                                $tel->setValue($infoDados->telefone);

                                $divL = new DivElement();
                                $divL->InsertRow('6',$nome);
                                $divL->InsertRow('6',$cpf);
                                $divL->InsertRow('6',$email);
                                $divL->InsertRow('6',$tel);
                                
                                
                                $divL->show();

                                $response["status"] = 200;
                                $response["header"] = "<b>Editar Perfil ".$infoDados->nomeCompleto."</b>";
                                $response["footer"] = '<button type="button" class="btn btn-success" data-dismiss="salvaForm" data-typeForm = "edit-user" data-infoId = "'.$infoDados->id.'">Salvar</button> <button type="button" class="btn btn-default" data-bs-dismiss="modal">Sair</button>';
                                $response["body"] = ob_get_contents(); // Store buffer in variable
                            }
                            else{
                                throw new Exception("Não foi possível encontrar os dados do usuário");
                            }
                        } catch (\Throwable $th) {
                            $response["status"] = 500;
                            $response["message"] = $th->getMessage();
                        }
                    }

                    ob_end_clean(); // End buffering and clean up
                    echo json_encode($response);
                    exit();
                break;

                case 3:
                    $cpf = Utility::getVariable("cpf");
                    
                    $response = InfoDadoss::search(json_encode(array("CPF"=>$cpf)));
                    $response = json_decode($response);
                    if($response->status==200){
                        $infoDadosPerito = new InfoDadoss($response->data[count($response->data)-1]);
                        $response = array();
                        $response['status'] = 200;
                        $response['cpf'] = $cpf;
                        $response['nome'] = $infoDadosPerito->nomeCompleto;
                    }
                    echo json_encode($response);
                    exit();
                break;
                
			}
		break;
        default:
            $div = new divElement();
            $hidden = "<input style='display:none' id='auxStade'>";
            $button = '<button type="button" id="newPersonal" class="btn btn-primary btn-lg btn-block">
                            <i class="mdi mdi-account"></i>                      
                            Criar novo Funcionário Personal
                        </button>';
            $div->InsertRow('4',$button);
            $button = '<button type="button" id="newFuncionario" class="btn btn-success btn-lg btn-block">
                            <i class="mdi mdi-account"></i>                      
                            Criar novo Funcionário Administrativo
                        </button>';
            $div->InsertRow('4',$button);
            $button = '<button type="button" id="newCliente" class="btn btn-dark btn-lg btn-block">
                            <i class="mdi mdi-account"></i>                      
                            Criar novo Cliente
                        </button>';
            $div->InsertRow('4',$button);
            $div->InsertRow('1',$hidden);
            $div->show();
            echo '<br>';
            $arrayBlock = array();
            $lines = array();
            $block = new GridBlock(""); 
            array_push($arrayBlock,$block);
            $block = new GridBlock("Nome");
            array_push($arrayBlock,$block);
            $block = new GridBlock("CPF"); 
            array_push($arrayBlock,$block);
            $block = new GridBlock("Email"); 
            array_push($arrayBlock,$block);
            $block = new GridBlock("Telefone"); 
            array_push($arrayBlock,$block);
            $block = new GridBlock("Acesso"); 
            array_push($arrayBlock,$block);
            $block = new GridBlock("Situação"); 
            array_push($arrayBlock,$block);

            $line = new GridLine($arrayBlock);
            $line->setHeader(true);
            array_push($lines,$line);
            $div = new DivElement();
            $grid = new GridElement($lines);
            $grid->setID("dataTables");
            $grid->setAjax($host."/pages/usuarios.php?status=ajax&case=1");
            $div->InsertRow('12',$grid);
            $div->show();
        break;
    }
Utility::InsertJavascriptFile("usuarios");
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