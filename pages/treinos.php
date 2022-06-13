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
                    // $text = new HeadingElement("","5","<b>Adicionar novo pagamento:</b>");
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
                    //$search = json_decode(ClientePerito::search(json_encode(array("id"=>$hash)),false,'like')); //mudar pra hash depois
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

                    $html = '<form id="submitPagamento" enctype="multipart/form-data" class="form-inline" style=" width: 100%; margin-top:10px" action="./pages/modalidades.php?status=saveTurma" method="POST"">

                                <div class="col-lg-4">
                                    <div style="" clas ns="form-group  ">
                                            <label style="" class="control-label ">Selecione a Modalidade da Turma</label>
                                            <input input-type="text" id="idModalidade" name="idModalidade" class="form-control " placeholder="Modalidade" value="" required="required" style="">
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
                                            <label style="" class="control-label ">Selecione Funcionário Responsável da Turma</label>
                                            <input input-type="text" id="idFuncionarioResp" name="idFuncionarioResp" class="form-control " placeholder="Nome da Modalidade" value="" required="required" style="">
                                    </div>
                                </div>
                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                            <label style="" class="control-label ">Quantidade de Vagas da Modalidade</label>
                                            <input type="number" id="quantidadeVagas" name="quantidadeVagas" class="form-control " placeholder="Quantidade de vagas" required="required" >
                                    </div>
                                </div>    
                                <div class="col-lg-5 align-self-end">
                                    <div class="form-group">
                                        <button type="submit" class="btn btn-success btn-lg" data-dismiss="salvaForm">Adicionar Modalidade</button>
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
            $idFuncionarioResp = Utility::getVariable("idFuncionarioResp");

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

                        $card = new Cards();
                        $card->setTextMain($modalidade->nome);
                        $card->setTextSub(" <b>Começa as : </b>".Utility::dateFormatToBR($turma->Hora_Inicio)."
                                            <b>Termina as: </b>".Utility::dateFormatToBR($turma->Hora_Final)."
                                            <br><b>Treinador Responsável:</b> "."Eduardo Atene Silva");//$allTurmas->idFuncionarioResponsavel);
                        $card->setTextButton("Ver Informações");
                        $card->setTextLeftButtonExist(true);
                        $card->setTextLeftButton("0/".$turma->Vagas);
                        $card->setImageExiste(false);
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

                        $card = new Cards();
                        $card->setTextMain($modalidade->nome);
                        $card->setTextSub("<b>Treinador Responsável:</b> "."Eduardo Atene Silva");//$allTurmas->idFuncionarioResponsavel);
                        $card->setTextButton("Ver Detalhes");
                        $card->setTextLeftButtonExist(false);
                        $card->setImageExiste(false);
                        $div1->InsertRow('3',$card);
                    }
                }
                else{
                    $allTurmas = array();
                }
                $div1->Show();
    
                Utility::InsertJavascriptFile("treinos");

            break;
            case 3: // 
                $botaoGenerateReport = new HtmlElement('<button type="button" id="relatorio-consult" class="btn btn-dark btn-lg btn-block"> <i class="mdi mdi-download btn-icon-prepend"></i> Gerar relatório</button>');
                $div = new divElement();
                $cardElements ='<div class="col-12  d-flex" >
                                    <div class="col-4">
                                        <button style="margin-right: 7px;" type="button" id="add-turma" class="btn btn-success btn-lg btn-block"> <i class="mdi mdi-plus btn-icon-prepend"></i> Cadastrar Treino</button>
                                    </div>
                                </div>';
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
                        $card->setTextButton("Ver Informações");
                        $card->setTextLeftButtonExist(true);
                        $card->setImageExiste(false);
                        $card->setTextLeftButton("0/".$turma->Vagas);
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
                Utility::InsertJavascriptFile("treinos");

            break;
        }
        break;
    }

    function getHistoryConsultoria($idProcesso_Clienteperito_id,$border = true){
        if($idProcesso_Clienteperito_id != null){
            $historicoConsultoria = json_decode(Consultoria::historicoConsultoria($idProcesso_Clienteperito_id));//$consultoria->tbprocesso_tbclienteperito_id));
            if(count($historicoConsultoria) != 0){
                $divConsultValues = new DivElement();
                $divConsultValues->setClass("p-3");
                $recente = true;
                foreach($historicoConsultoria as $hConsultoria){
                    $tempIndicada = new StatusTemperatura($hConsultoria->temperaturaNova);
                    if($recente){
                        $divConsultValues->setClass("p-3");
                        $divConsultValues->InsertRow("12","<b>Solicitada Consultoria dia ".Utility::dateFormatToBR($hConsultoria->dateEnvioConsultoria)."</b>");
                        $divConsultValues->InsertRow("12","<b>Análise: ".Utility::dateFormatToBR($hConsultoria->dataCheckedAdmin)."</b>");
                        $divConsultValues->InsertRow("12","<b>Temperatura Indicada: ".$tempIndicada->temperatura."</b>");
                        $divConsultValues->InsertRow("12","<b>Observações: ".$hConsultoria->comentariosAdmin."</b>");
                        if($hConsultoria->fileName != null )
                            $divConsultValues->InsertRow("12","<b>Petição indicada: </b> <a href='".Consultoria::getLinkFile($hConsultoria->hash)."' target='_blank'>link</a>");
                        else
                            $divConsultValues->InsertRow("12","<b>Petição indicada: Não indicada</b>");

                        $divConsultValues->InsertRow("12", new SeparatorElement());
                        $recente = false;

                        
                    }else{
                        $divConsultValues->setClass("p-3");
                        $divConsultValues->InsertRow("12","Solicitada Consultoria dia ".Utility::dateFormatToBR($hConsultoria->dateEnvioConsultoria)."</b>");
                        $divConsultValues->InsertRow("12","Análise: ".Utility::dateFormatToBR($hConsultoria->dataCheckedAdmin)."");
                        $divConsultValues->InsertRow("12","Temperatura Indicada: ".$tempIndicada->temperatura."");
                        $divConsultValues->InsertRow("12","Observações: ".$hConsultoria->comentariosAdmin."");
                        if($hConsultoria->fileName != null )
                            $divConsultValues->InsertRow("12","Petição indicada: <a href='".Consultoria::getLinkFile($hConsultoria->hash)."' target='_blank'>link</a>");
                        else
                            $divConsultValues->InsertRow("12","Petição indicada: Não indicada");
                        $divConsultValues->InsertRow("12", new SeparatorElement());
                    }

                }

                $divConsultHist = new DivElement();
                if($border == true){
                    $divConsultHist->setClass("border border-dark h-100");
                    $divConsultHist->setStyle("overflow-y:scroll");
                }
                $divConsultHist->InsertRow("12","<h3 class='text-center m-4' style='color:rgb(70,77,238); margin-top:8px' >Informações da Consultoria</h3>");
                $divConsultHist->InsertRow("12",$divConsultValues);
                return $divConsultHist;
            }else
                return null;
        }else
            return null;
    }
?>
