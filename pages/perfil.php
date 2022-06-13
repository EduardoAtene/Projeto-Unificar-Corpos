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
                    foreach ($_POST as $key => $value){
                        $data[$key] = $value;
                    }

                    if(array_key_exists("usuario",$data) && array_key_exists("senha",$data) && array_key_exists("tipoUsuario",$data)){ //cadastro de novo usuario
                        $response = Usuarioo::search(json_encode(array("login"=>$data["usuario"])));
                        $response = json_decode($response);
                        if($data["tipoUsuario"]=="funcionario")
                            $tbTipoAcesso_id = 1;
                        else if($data["tipoUsuario"]=="cliente")
                            $tbTipoAcesso_id = 2;
                        else if($data["tipoUsuario"]=="personal")
                            $tbTipoAcesso_id = 3;
                        if($response->status!=200){ //nao tem gente com esse login, seguro proceder
                            if($tbTipoAcesso_id==2){
                                $response = InfoDadoss::search(json_encode(array("CPF"=>$data["cpf"])));
                                $response = json_decode($response);
                                if($response->status==200){
                                    $dadosSave = array("email"=>$data["email"],"telefone"=>$data["tel"]);
                                    InfoDadoss::update($response->data[count($response->data)-1],json_encode($dadosSave));
                                    $response = Usuarioo::search(json_encode(array("idTipoAcesso_id"=>2,"idInfoDados"=>$response->data[count($response->data)-1])));
                                    $response = json_decode($response);
                                    if($response->status==200){
                                        $dados = array("login"=>$data["usuario"],"senha"=>sha1("kito".$data["senha"]));
                                        Usuarioo::update($response->data[count($response->data)-1],json_encode($dados));
                                    }
                                }
                                else{
                                    $dados = array("nomeCompleto"=>$data["nomeCompleto"],"CPF"=>$data["cpf"],"email"=>$data["email"],"telefone"=>$data["tel"]);
                                    $idInfoDados = InfoDadoss::set(json_encode($dados));
                                    $dados = array("login"=>$data["usuario"],"senha"=>sha1("kito".$data["senha"]),"idTipoAcesso_id"=>$tbTipoAcesso_id,"idInfoDados"=>$idInfoDados);
                                    $idUsuario = Usuarioo::set(json_encode($dados));
                                    $dados = array("idUsuario"=>$idUsuario, "dataInicio" => date("Y-m-d H:i:s"));
                                    $idNewCliente = Cliente::set(json_encode($dados));
                                }
                            }
                            else{
                                $dados = array("nomeCompleto"=>$data["nomeCompleto"],"CPF"=>$data["cpf"],"email"=>$data["email"],"telefone"=>$data["tel"]);
                                $idInfoDados = InfoDadoss::set(json_encode($dados));
                                $dados = array("login"=>$data["usuario"],"senha"=>sha1("kito".$data["senha"]),"idTipoAcesso_id"=>$tbTipoAcesso_id,"idInfoDados"=>$idInfoDados);
                                $idUsuario = Usuarioo::set(json_encode($dados));
                                $dados = array("idUsuario"=>$idUsuario, "cargo" => $data["cargo"]);
                                $idNewCliente = Funcionario::set(json_encode($dados));
                            }
                        }
                        else{
                            echo -1;
                        }
                    }
                    else if(array_key_exists("tipoUsuario",$data) && array_key_exists("infoId",$data)){ //alteração no perfil de outro usuario
                        try {
                            if($data["infoId"]>0){
                                InfoDadoss::update($data["infoId"],json_encode(array("nomeCompleto"=>$data["nomeCompleto"], "CPF"=>$data["cpf"],"email"=>$data["email"],"telefone"=>$data["tel"])));
                                $response["status"] = 200;
                                $response["message"] = 'success';
                            }
                            else{
                                throw new Exception("Não foi possível salvar os dados do usuario");
                            }
                        } catch (\Throwable $th) {
                            $response["status"] = 500;
                            $response["message"] = $th->getMessage();
                        }
                    }
                    else{ 
                        $user = new Usuarioo($_SESSION["idUser"]);
                        $password = Utility::getVariable("pass");

                        $search = array("login"=>$user->login,"pass"=>$password);
                        $response = Usuarioo::tryLogin($search);
                        if($response != -1){           
                                $infoDados = new InfoDadoss($user->idInfoDados);
                                $_SESSION["nome"] = $infoDados->nomeCompleto;
                                $infoDados->email = $data["email"];
                                $infoDados->telefone = $data["tel"];
            
                                $infoDados->save();
                                $retorno = array();
                                $retorno['status']= 200;
                                $retorno['passwordValid']= true;
                                echo json_encode($retorno);

                        }else{
                            $retorno = array();
                            $retorno['status']= 200;
                            $retorno['passwordValid']= false;
                            echo json_encode($retorno);
                        }
  
                    }

                    exit();
                break;
                case 2:
                    foreach ($_POST as $key => $value){
                        $data[$key] = $value;
                    }
                    $user = new Usuarioo($_SESSION["idUser"]);
                    $password = $data["pass"];

                    $search = array("login"=>$user->login,"pass"=>$password);
                    $response = Usuarioo::tryLogin($search);
                    if($response != -1){
                        $clientePerito->save();
                        $retorno = array();
                        $retorno['status']= 200;
                        $retorno['passwordValid']= true;
                        echo json_encode($retorno);
                    }else{
                        $retorno = array();
                        $retorno['status']= 200;
                        $retorno['passwordValid']= false;
                        echo json_encode($retorno);
                    }
                    
                break;
                case 3:
                    $user = new Usuarioo($_SESSION["idUser"]);
                    $response = array();
                    if(count($_FILES) > 0){
                        $path = $_FILES['files']['tmp_name'];
                        $type = pathinfo($path, PATHINFO_EXTENSION);
                        $data = file_get_contents($path);
                        $base64 = base64_encode($data);
                        $file = $base64;
                        $_SESSION["file"] = $base64;
                        Usuarioo::update($user->id,json_encode(array("img" => $base64)));
                        $response["status"] = 200;
                        $response["imgSession"] = "data:image/png;base64,".$base64;
                        $response["message"] = 'Imagem Inserida com Sucesso!';
                        echo json_encode($response);
                    }else{
                        $response["status"] = 401;
                        $response["message"] = 'Imagem Não Inserida!';
                        echo json_encode($response);
                    }
                break;
			}
		break;
        default:
            $divLoadingUpload = new DivElement();
            $divLoadingUpload->setID("container");
            $divLoadingUpload->setVisible(false);
            $divLoadingUpload->show();
            if(isset($session)){
                $user = new Usuarioo($session->id);
                $infoDados = new InfoDadoss($user->idInfoDados);

            }
            isset($infoDados)? $nomePerfil = " de ".$infoDados->nomeCompleto : $nomePerfil = "";
            $header = HeaderElement::HeaderDefault("Perfil".$nomePerfil);
			$nome = new InputElement("nome","Nome Completo","text","","seu nome","250");
            if(isset($infoDados)){
                $nome->setValue($infoDados->nomeCompleto);
                $nome->setActive(false);
            }
            $cpf =new InputElement("cpf","CPF","number","","digite somente números");
            if(isset($infoDados)){
                $cpf->setValue($infoDados->CPF);
                $cpf->setActive(false);
            }
            $email = new InputElement("email","Email","email","","fulano@algo.com");
            if(isset($infoDados))
                $email->setValue($infoDados->email);
            $tel = new InputElement("tel","Telefone de Contato Celular","number","","DDD + Número");
            if(isset($infoDados))
                $tel->setValue($infoDados->telefone);

            $dataNascimento = new InputElement("dataNascimento","Data Nascimento","text","","Data Nascimento");
            if(isset($infoDados)){
                $dataNascimento->setValue($infoDados->dataNascimento);
                $dataNascimento->setActive(false);
            }
            $labelImage = new HTMLElement('<b><label>Foto de Perfil<label/></b>');
            if(isset($user) && $user->img!=''){
                $image = new HTMLElement('<input id = "my-img" type="image" src="data:image/png;base64, '.$user->img.'" width="150px";height:"150px" style="border: #6c7293 solid 2px;border-radius: 10px;margin-top:10px"/>
                <input type="file" id="img" accept="image/jpeg, image/png" style="display: none;" />');
            }else{
                $image = new HTMLElement('<input id = "my-img" type="image"  src="'.$host.'/images/user_default.png" width="150px";height:"150px" style="border: #6c7293 solid 2px;border-radius: 10px;margin-top:10px"/>
                <input type="file" id="img" accept="image/jpeg,image/jpg, image/png" style="display: none;" />');
            }

            $salvaAlteracoes = new buttonElement("salvaAlteracoes","Salvar Alterações");
            $salvaAlteracoes->setClass("btn btn-primary btn-rounded btn-fw btn-lg");

            $divInfo = new DivElement();
            $divInfo->InsertRow('12',$nome);
            $divInfo->InsertRow('8',$cpf);
            $divInfo->InsertRow('4',$dataNascimento);

            $divImage = new DivElement();
            $divImage->setClass("text-center");    
            $divImage->InsertRow('12',$labelImage);

            if(isset($image))
                $divImage->InsertRow('12',$image);
                
            $divL = new DivElement();
            $divL->InsertRow('6',$divInfo);
            $divL->InsertRow('6',$divImage);
            $divL->InsertRow('6',$email);
            $divL->InsertRow('6',$tel);


            $divMain = new DivElement();
            $divMain->InsertRow('12',$divL);
            $divMain->InsertRow('12 text-center',$salvaAlteracoes);
            
            $divMain->show();
            Utility::InsertJavascriptFile("perfil");
            ?>
            
            <div class="modal fade" id="modal" role="dialog" style="height: 100vh">
                <div class="modal-dialog modal-lg" style = "width:100%;max-width:1250px">
                    <div class="modal-content" style="margin-top: -3%;">
                        <div class="modal-header" style="position:relative; top: 0; left: 0; right: 17px; z-index: 10; background: white;">
                            <button type="button" class="close" data-dismiss="modal">&times;</button>
                            <h4 class="modal-title"></h4>
                        </div>
                        <div style="overflow-y: scroll; height: 80%;" class="modal-body">						
                                                                                        
                        </div>
                        <div class="modal-footer" style="position:relative; bottom: 0; left: 0; right: 17px; z-index: 10; background: white;">
                            <button type="button" class="btn btn-success" data-dismiss="salvaForm">Salvar</button>
                            <button type="button" class="btn btn-default" data-dismiss="modal">Sair</button>
                        </div>
                    </div>
                </div>
            </div>
            <?php
        break;
    }
?>

