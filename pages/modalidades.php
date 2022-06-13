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
                case 1: // Modal da Modalidades
                    ob_start(); // Start output buffering
                    $header = HeaderElement::HeaderDefault("Criar Modalidade");

                    $html = '<form id="submitPagamento" enctype="multipart/form-data" class="form-inline" style=" width: 100%;" action="./pages/modalidades.php?status=saveModalidade" method="POST"">

                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                            <label style="" class="control-label ">Nome da Modalidade a ser Adicionada</label>
                                            <input input-type="text" id="nomeModalidade" maxlength="40" name="nomeModalidade" class="form-control text " placeholder="Nome da Modalidade" value="" required="required" style="">
                                    </div>
                                </div>
        
                                <div class="col-lg-4">
                                    <div style="" class="form-group">
                                        <label style="" class="control-label">Imagem da Modalidade para apresentação</label>
                                        <div class="input-group">
                                            <input type="file" id="file" name="file" required="required" accept="image/jpeg, image/png" class="form-control" id="inputGroupFile04" aria-describedby="inputGroupFileAddon04" aria-label="Upload">
                                        </div>
                                    </div>

                                </div>

                                <div class="col-lg-4">
                                    <div class="form-group">
                                        <button type="submit" class="btn btn-success btn-lg" data-dismiss="salvaForm">Adicionar Modalidade</button>
                                    </div>
                                </div>
							</form>';

                    $colunas = array("Nome da Modalidade","Imagem da Modalidade");
                    $ajaxPath = $host."/pages/modalidades.php?status=ajax&case=2";
                    $grid = GridElement::generatePrototype($colunas,$ajaxPath);

                    $div = new DivElement();
                    $divInline = new DivElement();
                    $divInline->setClass('form-inline');
                    $div->InsertRow('12',$header);
                    // $div->InsertRow('12',$text);
                    $div->show();
                    echo $html;
                    $header = HeaderElement::HeaderDefault("Modalidades");

                    $div = new DivElement();
                    $div->InsertRow('12',$grid);
                    $div->show();
                    try {
                        if(true){
                            $response["status"] = 200;
                            $response["header"] = "<b>Todas Modalidades do Sistema</b>";
                            $response["body"] = ob_get_contents(); // Store buffer in variable
                            $response["footer"] = '<button type="button" class="btn btn-default" data-bs-dismiss="modal">Sair</button>';
                        }
                        else{
                            throw new Exception("Não foi possível encontrar os dados do processo");
                        }
                    } catch (\Throwable $th) {
                        $response["status"] = 500;
                        $response["message"] = $th->getMessage();
                    }
                    ob_end_clean(); // End buffering and clean up
                    echo json_encode($response);
                    exit();
                break;
                case 2: // Ajax da Tabela de Modalidades
                    // $search = json_decode(PagamentoCliente::GetAll());
                    $modalidades = Modalidade::GetAll();
                    $temp = array();
                        foreach($modalidades as $idModalidade){
                            $modalidade = new Modalidade($idModalidade);
                            if($modalidade->image!=NULL)
                                $botaoDownload = '  <button name = "downloadImagem" type="button" idModalidade = "'.$modalidade->id.'" class="btn btn-outline-success btn-icon-text">
                                                        <i class="mdi mdi-download btn-icon-prepend"></i>                   
                                                        Download
                                                    </button>';
                            else
                                $botaoDownload = 'Não possui Imagem';

                            array_push($temp,array(
                                $modalidade->nome,
                                $botaoDownload,

                            ));
                        }           
                    if(count($temp)>0){
                        $retorno["status"] = 200;
                        $retorno["data"] = $temp;
                    }
                    else{
                        $retorno["status"] = 500;
                        $retorno["data"] = array();
                    }
                    echo json_encode($retorno);
                    exit();
                break;
                break;
                case 3: // Download Imagem da Modalidade
                    $idModalidade = Utility::getvariable("idModalidade");
                    $modalidade = new Modalidade($idModalidade);

                    $base64 = $modalidade->image;
                    $decoded = base64_decode($base64);
                    $file = dirname(__DIR__)."/files/temp/".$modalidade->imageNameFile;
                    file_put_contents($file, $decoded);
                    
                    if (file_exists($file)) {
                        header('Content-Type: image/png');
                        header('Content-Disposition: attachment; filename="'.basename($file).'"');
                        header('Content-Length: ' . filesize($file));

                        readfile($file);
                        // unlink($file);
                        exit();
                    }
                    exit(); 
                break;  
                case 4: // Modal da Modalidades
                    ob_start(); // Start output buffering
                    $header = HeaderElement::HeaderDefault("Criar Turma");

                    $valuesModalidade = $valuesPersonal = "";

                    $modalidades = Modalidade::GetAll();
                    foreach($modalidades as $idModalidade){
                        $modalidade = new Modalidade($idModalidade);
                        $valuesModalidade .= "<option value='".$idModalidade."'>".$modalidade->nome."</option>";
                    }

                    $personais = Funcionario::getAllPersonal();
                    foreach($personais as $idPersonal){
                        $infoPersonal = Funcionario::getInfoDados($idPersonal);
                        $valuesPersonal .= "<option value='".$idPersonal."'>".$infoPersonal['nomeCompleto']."</option>";
                    }
                    $html = '<form id="submitPagamento" enctype="multipart/form-data" class="form-inline" style=" width: 100%; margin-top:10px" action="./pages/modalidades.php?status=saveTurma" method="POST"">

                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                            <label style="" class="control-label ">Modalidade da Turma</label>
                                            <select required="required" id="modalidade" class="form-select" name="modalidade" aria-label="Selecione a Modalidade da Turma">
                                                <option value="-1" disabled selected>Selecione a Modalidade da Turma</option>
                                                '.$valuesModalidade.'
                                            </select>
                                            <input style="display:none" type="number" id="idModalidade" name="idModalidade" class="form-control " placeholder="Quantidade de vagas" required="required">

                                    </div>
                                </div>
                                
                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                        <label style="" class="control-label ">Data/Hora Inicio da Modalidade</label>
                                        <input  class="form-control" type="datetime-local" id="data_inicio"  name="data_inicio" required="required">
                                    </div>
                                </div>
                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                        <label style="" class="control-label ">Data/Hora Fim da Modalidade</label>
                                        <input  class="form-control" type="datetime-local" id="data_fim"  name="data_fim" required="required">
                                    </div>
                                </div>
                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                            <label style="" class="control-label ">Funcionário Responsável</label>
                                            <select required="required" id="funcionarioResp" class="form-select"  name="funcionarioResp" aria-label="Selecione a Modalidade da Turma">
                                                <option value="-1" disabled selected>Selecione Funcionário Responsável da Turma</option>
                                                '.$valuesPersonal.'
                                            </select>
                                            <input style="display:none" type="number" id="idFuncionarioResp" name="idFuncionarioResp" class="form-control " placeholder="Quantidade de vagas"  required="required">

                                    </div>
                                </div>
                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                            <label style="" class="control-label ">Quantidade de Vagas da Modalidade</label>
                                            <input type="number" id="quantidadeVagas" name="quantidadeVagas" class="form-control " placeholder="Quantidade de vagas"  >
                                    </div>
                                </div>    
                                <div class="col-lg-5 align-self-end">
                                    <div class="form-group">
                                        <button id="buttonSaveTurma" type="submit" class="btn btn-success btn-lg" data-dismiss="salvaForm">Adicionar Modalidade</button>
                                    </div>
                                </div>
							</form>';

                    $div = new DivElement();
                    $divInline = new DivElement();
                    $divInline->setClass('form-inline');
                    $div->InsertRow('12',$header);
                    $div->show();
                    echo $html;
                    try {
                        if(true){
                            $response["status"] = 200;
                            $response["header"] = "<b>Criar Turmas no Sistema de Academia</b>";
                            $response["body"] = ob_get_contents(); // Store buffer in variable
                            $response["footer"] = '<button type="button" class="btn btn-default" data-bs-dismiss="modal">Sair</button>';
                        }
                        else{
                            throw new Exception("Não foi possível encontrar os dados do processo");
                        }
                    } catch (\Throwable $th) {
                        $response["status"] = 500;
                        $response["message"] = $th->getMessage();
                    }
                    ob_end_clean(); // End buffering and clean up
                    echo json_encode($response);
                    exit();
                break;
                case 5: // Inscrever / Desinscrever da Turma
                    $idTurma = Utility::getVariable("idTurma");
                    $desinscrever = Utility::getVariable("desinscrever");             

                    $response = array();
                    $response['status'] = 400;
                    if($desinscrever == "false"){
                        $idUsuario = $_SESSION['id'];
                        $idCliente = Cliente::getIdByIdUsuario($idUsuario);

                        $save = array();
                        $save['idTurma'] = $idTurma;
                        $save['idCliente'] = $idCliente;
                        Turma_Cliente::set(json_encode($save));

                        $response['status'] = 200;
                    }else{
                        $idTurmaCliente = Turma_Cliente::userIsExistInTurma($idTurma, Cliente::getIdByIdUsuario($_SESSION['id']));

                        $turma_cliente_delete = new Turma_Cliente($idTurmaCliente);
                        $turma_cliente_delete->Delete();

                        $response['status'] = 200;

                    }

                    echo json_encode($response);
                    exit();
                break;
                case 6: // Modal de ver detalhes da modalidade
                    ob_start(); // Start output buffering
                    $header = HeaderElement::HeaderDefault("Criar Modalidade");

                    $html = '<form id="submitPagamento" enctype="multipart/form-data" class="form-inline" style=" width: 100%;" action="./pages/modalidades.php?status=saveModalidade" method="POST"">

                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                            <label style="" class="control-label ">Nome da Modalidade a ser Adicionada</label>
                                            <input input-type="text" id="nomeModalidade" maxlength="40" name="nomeModalidade" class="form-control text " placeholder="Nome da Modalidade" value="" required="required" style="">
                                    </div>
                                </div>
        
                                <div class="col-lg-4">
                                    <div style="" class="form-group">
                                        <label style="" class="control-label">Imagem da Modalidade para apresentação</label>
                                        <div class="input-group">
                                            <input type="file" id="file" name="file" required="required" accept="image/jpeg, image/png" class="form-control" id="inputGroupFile04" aria-describedby="inputGroupFileAddon04" aria-label="Upload">
                                        </div>
                                    </div>

                                </div>

                                <div class="col-lg-4">
                                    <div class="form-group">
                                        <button type="submit" class="btn btn-success btn-lg" data-dismiss="salvaForm">Adicionar Modalidade</button>
                                    </div>
                                </div>
							</form>';

                    $colunas = array("Nome da Modalidade","Imagem da Modalidade");
                    $ajaxPath = $host."/pages/modalidades.php?status=ajax&case=2";
                    $grid = GridElement::generatePrototype($colunas,$ajaxPath);

                    $div = new DivElement();
                    $divInline = new DivElement();
                    $divInline->setClass('form-inline');
                    $div->InsertRow('12',$header);
                    // $div->InsertRow('12',$text);
                    $div->show();
                    echo $html;
                    $header = HeaderElement::HeaderDefault("Modalidades");

                    $div = new DivElement();
                    $div->InsertRow('12',$grid);
                    $div->show();
                    try {
                        if(true){
                            $response["status"] = 200;
                            $response["header"] = "<b>Todas Modalidades do Sistema</b>";
                            $response["body"] = ob_get_contents(); // Store buffer in variable
                            $response["footer"] = '<button type="button" class="btn btn-default" data-bs-dismiss="modal">Sair</button>';
                        }
                        else{
                            throw new Exception("Não foi possível encontrar os dados do processo");
                        }
                    } catch (\Throwable $th) {
                        $response["status"] = 500;
                        $response["message"] = $th->getMessage();
                    }
                    ob_end_clean(); // End buffering and clean up
                    echo json_encode($response);
                    exit();
                break;
			}
		break;
        
        case 'saveModalidade':
            $nomeModalidade = Utility::getVariable("nomeModalidade");
            $arraySave['nome']=$nomeModalidade;

            if(count($_FILES) > 0){
                $path = $_FILES['file']['tmp_name'];
                $name = $_FILES['file']['name'];
                $type = pathinfo($path, PATHINFO_EXTENSION);
                $data = file_get_contents($path);
                $base64 = base64_encode($data);
                $file = $base64;
                $arraySave['image']=$file;
                $arraySave['imageNameFile']=$name;
            }
            Modalidade::set(json_encode($arraySave));
            header("Location: ".$_SERVER['HTTP_ORIGIN']."/modalidades"); 
        break;
                
        case 'saveTurma':
            $idModalidade = Utility::getVariable("idModalidade");
            $data_inicio = Utility::getVariable("data_inicio");
            $data_fim = Utility::getVariable("data_fim");
            $quantidadeVagas = Utility::getVariable("quantidadeVagas");
            $idFuncionarioResp = Utility::getVariable("funcionarioResp");

            $arraySave['idModalidade']=$idModalidade;
            $arraySave['Hora_Inicio']= Utility::dateFormatToEU($data_inicio);
            $arraySave['Hora_Final']= Utility::dateFormatToEU($data_fim);
            $arraySave['Vagas']=$quantidadeVagas;
            $arraySave['idFuncionarioResponsavel']=$idFuncionarioResp;

            Turma::set(json_encode($arraySave));
            header("Location: ".$_SERVER['HTTP_ORIGIN']."/modalidades"); 
        break;
        default:
        $json = json_encode(array("id"=>$session->id));
        $user = new Usuarioo($session->id);
        $tipoAcesso = $user->idTipoAcesso_id;
        switch ($tipoAcesso){
            case 1: // 
                $botaoGenerateReport = new HtmlElement('<button type="button" id="relatorio-consult" class="btn btn-dark btn-lg btn-block"> <i class="mdi mdi-download btn-icon-prepend"></i> Gerar relatório</button>');
                $div = new divElement();
                $cardElements ='<div class="col-12  d-flex" >
                                    <div class="col-4">
                                        <button style="margin-right: 7px;" type="button" id="add-turma" class="btn btn-success btn-lg btn-block"> <i class="mdi mdi-plus btn-icon-prepend"></i> Criar Turma</button>
                                    </div>
                                    <div class="col-4">
                                        <button type="button" id="add-modalidade" class="btn btn-primary btn-lg btn-block"> <i class="mdi mdi-plus btn-icon-prepend"></i> Criar Modalidade</button>
                                </div>
                                </div>';
                $allTurmas = json_decode(Turma::get());
                $div1 = new DivElement();
                if($allTurmas->status =='200'){
                    foreach($allTurmas->data as $turma){
                        $modalidade = new Modalidade($turma->idModalidade);
                        $infoFuncionario = Funcionario::getInfoDados($turma->idFuncionarioResponsavel);
                        $quantidadeInscritos = Turma_Cliente::getQuantidadeInscritos($turma->id);
                        $card = new Cards();
                        $card->setTextMain($modalidade->nome);
                        $card->setTextSub(" <b>Começa as : </b>".Utility::dateFormatToBR($turma->Hora_Inicio)."
                                            <b>Termina as: </b>".Utility::dateFormatToBR($turma->Hora_Final)."
                                            <br><b>Treinador Responsável:</b> ".$infoFuncionario["nomeCompleto"]);//$allTurmas->idFuncionarioResponsavel);
                        $card->setTextButton("Ver Informações");
                        $card->setTextLeftButtonExist(true);
                        $card->setTextLeftButton($quantidadeInscritos."/".$turma->Vagas);
                        $card->setImageExiste(true);
                        $card->setImagemBase($modalidade->image,$modalidade->imageNameFile);
                        $div1->InsertRow('3',$card);
                    }
                }
                else{
                    $allTurmas = array();
                }

                $div->InsertRow('12',$cardElements);
                $div->InsertRow('12',$div1);
                $div->show();
                ?>
                    <div class="modal fade" id="modalModalidade" role="dialog" style="height: 100vh">
                        <div class="modal-dialog modal-lg" style = "width:100%;max-width:90vw">
                            <div class="modal-content">
                                <div class="modal-header" style="position:relative; top: 0; left: 0; right: 17px; z-index: 10; background: white;">
                                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                                    <h4 class="modal-title"></h4>
                                </div>
                                <div style="overflow-y: scroll; height: 80%;" class="modal-body">						
                                                                                                
                                </div>
                                <div class="modal-footer" style="position:relative; bottom: 0; left: 0; right: 17px; z-index: 10; background: white;">

                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="modal fade" id="modalTurma" role="dialog" style="height: 100vh">
                        <div class="modal-dialog modal-lg" style = "width:100%;max-width:90vw">
                            <div class="modal-content">
                                <div class="modal-header" style="position:relative; top: 0; left: 0; right: 17px; z-index: 10; background: white;">
                                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                                    <h4 class="modal-title"></h4>
                                </div>
                                <div style="overflow-y: scroll; height: 80%;" class="modal-body">						
                                                                                                
                                </div>
                                <div class="modal-footer" style="position:relative; bottom: 0; left: 0; right: 17px; z-index: 10; background: white;">

                                </div>
                            </div>
                        </div>
                    </div>
                <?php
                Utility::InsertJavascriptFile("modalidades");

            break;
            case 2:
                $allTurmas = json_decode(Turma::get());
                $div1 = new DivElement();
                if($allTurmas->status =='200'){
                    foreach($allTurmas->data as $turma){
                        $modalidade = new Modalidade($turma->idModalidade);
                        $infoFuncionario = Funcionario::getInfoDados($turma->idFuncionarioResponsavel);

                        $userInscrito = Turma_Cliente::userIsExistInTurma($turma->id ,Cliente::getIdByIdUsuario($session->id));
                        $quantidadeInscritos = Turma_Cliente::getQuantidadeInscritos($turma->id);
                        $card = new Cards();
                        $card->setTextMain($modalidade->nome);
                        $card->setTextSub(" <b>Começa as : </b>".Utility::dateFormatToBR($turma->Hora_Inicio)."
                                            <b>Termina as: </b>".Utility::dateFormatToBR($turma->Hora_Final)."
                                            <br><b>Treinador:</b> ".$infoFuncionario["nomeCompleto"]);//$allTurmas->idFuncionarioResponsavel);
                 
                        if($userInscrito == -1){
                            if($quantidadeInscritos == $turma->Vagas){
                                $card->setDisable();
                                $card->setTextButton("Lotado");
                                $card->setColorButtonbyClass("btn-danger");
                                $card->setElementAddButton("input-desinscrever=false");
                            }else{
                                $card->setTextButton("Inscrever");
                                $card->setColorButtonbyClass("btn-primary");
                                $card->setElementAddButton("input-desinscrever=false");
                            }
                        }else{
                            $card->setTextButton("Desinscrever");
                            $card->setColorButtonbyClass("btn-danger");
                            $card->setElementAddButton("input-desinscrever=true");
                        }
                        $card->setIdButton($turma->id);
                        $card->setElementAddButton("input-name='".$modalidade->nome."'");
                        $card->setClass($turma->id);
                        $card->setTextLeftButtonExist(true);
                        $card->setTextLeftButton($quantidadeInscritos."/".$turma->Vagas);
                        $card->setImageExiste(true);
                        $card->setImagemBase($modalidade->image,$modalidade->imageNameFile);
                        $div1->InsertRow('3',$card);
                    }
                }
                else{
                    $header = HeaderElement::HeaderDefault("Nenhuma Turma Disponível");
                }
                $div1->Show();
    
                Utility::InsertJavascriptFile("modalidades");

            break;
            case 3:
                $allTurmas = json_decode(Turma::get());
                $div1 = new DivElement();
                if($allTurmas->status =='200'){
                    foreach($allTurmas->data as $turma){
                        $modalidade = new Modalidade($turma->idModalidade);

                        $card = new Cards();
                        $card->setTextMain($modalidade->nome);
                        $card->setTextSub(" <b>Começa as : </b>".Utility::dateFormatToBR($turma->Hora_Inicio)."
                                            <b>Termina as: </b>".Utility::dateFormatToBR($turma->Hora_Final)."
                                            <br><b>Treinador Responsável:</b> "."Eduardo Atene Silva");//$allTurmas->idFuncionarioResponsavel);
                        $card->setTextButton("Ver Detalhes");
                        $card->setTextLeftButtonExist(true);
                        $card->setTextLeftButton("0/".$turma->Vagas);
                        $card->setImageExiste(true);
                        $card->setImagemBase($modalidade->image,$modalidade->imageNameFile);
                        $div1->InsertRow('3',$card);
                    }
                }
                else{
                    $allTurmas = array();
                }

                $div1->Show();
    
                Utility::InsertJavascriptFile("modalidades");

            break;
        }
        break;
    }
?>
