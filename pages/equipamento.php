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
                    $header = HeaderElement::HeaderDefault("Criar Equipamento");

                    $html = '<form id="submitPagamento" enctype="multipart/form-data" class="form-inline" style=" width: 100%;" action="./pages/equipamento.php?status=saveEquipamento" method="POST"">

                                <div class="col-lg-4">
                                    <div style="" class="form-group  ">
                                            <label style="" class="control-label ">Nome do Equipamento a ser Adicionada</label>
                                            <input input-type="text" id="nomeModalidade" maxlength="40" name="nomeEquipamento" class="form-control text " placeholder="Nome do Equipamento" value="" required="required" style="">
                                    </div>
                                </div>
        
                                <div class="col-lg-4">
                                    <div style="" class="form-group">
                                        <label style="" class="control-label">Imagem do Equipamento para apresentação</label>
                                        <div class="input-group">
                                            <input type="file" id="file" name="file" required="required" accept="image/jpeg, image/png" class="form-control" id="inputGroupFile04" aria-describedby="inputGroupFileAddon04" aria-label="Upload">
                                        </div>
                                    </div>

                                </div>

                                <div class="col-lg-4">
                                    <div class="form-group">
                                        <button type="submit" class="btn btn-success btn-lg" data-dismiss="salvaForm">Adicionar Equipamento</button>
                                    </div>
                                </div>
							</form>';

                    $colunas = array("Nome do Equipamento","Imagem do Equipamento");
                    $ajaxPath = $host."/pages/equipamento.php?status=ajax&case=2";
                    $grid = GridElement::generatePrototype($colunas,$ajaxPath);

                    $div = new DivElement();
                    $divInline = new DivElement();
                    $divInline->setClass('form-inline');
                    $div->InsertRow('12',$header);
                    // $div->InsertRow('12',$text);
                    $div->show();
                    echo $html;
                    $header = HeaderElement::HeaderDefault("Equipamentos");

                    $div = new DivElement();
                    $div->InsertRow('12',$grid);
                    $div->show();
                    try {
                        if(true){
                            $response["status"] = 200;
                            $response["header"] = "<b>Equipamentos do Sistema</b>";
                            $response["body"] = ob_get_contents();
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
                    $equipamentos = Equipamento::GetAll();
                    $temp = array();
                        foreach($equipamentos as $idEquipamento){
                            $equipamento = new Equipamento($idEquipamento);
                            if($equipamento->image!=NULL)
                                $botaoDownload = '  <button name = "downloadImagem" type="button" idEquipamento = "'.$equipamento->id.'" class="btn btn-outline-success btn-icon-text">
                                                        <i class="mdi mdi-download btn-icon-prepend"></i>                   
                                                        Download
                                                    </button>';
                            else
                                $botaoDownload = 'Não possui Imagem';

                            array_push($temp,array(
                                $equipamento->nome,
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
                    $idEquipamento = Utility::getvariable("idEquipamento");
                    $equipamento = new Equipamento($idEquipamento);

                    $base64 = $equipamento->image;
                    $decoded = base64_decode($base64);
                    $file = dirname(__DIR__)."/files/temp/".$equipamento->imageNameFile;
                    file_put_contents($file, $decoded);
                    
                    if (file_exists($file)) {
                        header('Content-Type: image/png');
                        header('Content-Disposition: attachment; filename="'.basename($file).'"');
                        header('Content-Length: ' . filesize($file));

                        readfile($file);
                        unlink($file);
                        exit();
                    }
                    exit(); 
                break; 

			}
		break;
        
        case 'saveEquipamento':
            $nomeEquipamento = Utility::getVariable("nomeEquipamento");
            $arraySave['nome']=$nomeEquipamento;

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
            Equipamento::set(json_encode($arraySave));
            header("Location: ".$_SERVER['HTTP_ORIGIN']."/equipamento"); 
        break;

        default:
        $user = new Usuarioo($session->id);
        $tipoAcesso = $user->idTipoAcesso_id;
        
        switch ($tipoAcesso){
            case 1: // 
                $botaoGenerateReport = new HtmlElement('<button type="button" id="relatorio-consult" class="btn btn-dark btn-lg btn-block"> <i class="mdi mdi-download btn-icon-prepend"></i> Gerar relatório</button>');
                $div = new divElement();
                $cardElements ='<div class="col-12  d-flex" >
                                    <div class="col-4">
                                        <button style="margin-right: 7px;" type="button" id="add-equipamento" class="btn btn-success btn-lg btn-block"> <i class="mdi mdi-plus btn-icon-prepend"></i> Adicionar Equipamento</button>
                                    </div>
                                </div>';
                $allEquipamentos = Equipamento::GetAll();
                $div1 = new DivElement();
                foreach($allEquipamentos as $idEquipamento){
                    $equipamento = new Equipamento($idEquipamento);

                    $card = new Cards();
                    $card->setTextMain($equipamento->nome);
                    $card->setTextSub("<b>Treinador Responsável:</b> "."Eduardo Atene Silva");//$allTurmas->idFuncionarioResponsavel);
                    $card->setClassTextSub("d-none");
                    $card->setTextLeftButtonExist(false);;
                    $card->setImageExiste(true);
                    $card->setImagemBase($equipamento->image,$equipamento->imageNameFile);
                    $card->setColorCardbyClass("text-center");
                    $card->setColorButtonbyClass("d-none");
                    $div1->InsertRow('3',$card);
                }

                $div->InsertRow('12',$cardElements);
                $div->InsertRow('12',$div1);
                $div->show();
                ?>
                    <div class="modal fade" id="modalEquipamento" role="dialog" style="height: 100vh">
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
                Utility::InsertJavascriptFile("equipamento");

            break;
            case 2:
                $allEquipamentos = Equipamento::GetAll();
                $div1 = new DivElement();
                foreach($allEquipamentos as $idEquipamento){
                    $equipamento = new Equipamento($idEquipamento);

                    $card = new Cards();
                    $card->setTextMain($equipamento->nome);
                    $card->setTextSub("<b>Treinador Responsável:</b> "."Eduardo Atene Silva");//$allTurmas->idFuncionarioResponsavel);
                    $card->setClassTextSub("d-none");
                    $card->setTextLeftButtonExist(false);;
                    $card->setImageExiste(true);
                    $card->setImagemBase($equipamento->image,$equipamento->imageNameFile);
                    $card->setColorCardbyClass("text-center");
                    $card->setColorButtonbyClass("d-none");
                    $div1->InsertRow('3',$card);
                }

                $div1->show();
            break;
            case 3: // 
                $botaoGenerateReport = new HtmlElement('<button type="button" id="relatorio-consult" class="btn btn-dark btn-lg btn-block"> <i class="mdi mdi-download btn-icon-prepend"></i> Gerar relatório</button>');
                $div = new divElement();
                $cardElements ='<div class="col-12  d-flex" >
                                    <div class="col-4">
                                        <button style="margin-right: 7px;" type="button" id="add-equipamento" class="btn btn-success btn-lg btn-block"> <i class="mdi mdi-plus btn-icon-prepend"></i> Adicionar Equipamento</button>
                                    </div>
                                </div>';
                $allEquipamentos = Equipamento::GetAll();
                $div1 = new DivElement();
                foreach($allEquipamentos as $idEquipamento){
                    $equipamento = new Equipamento($idEquipamento);

                    $card = new Cards();
                    $card->setTextMain($equipamento->nome);
                    $card->setTextSub("<b>Treinador Responsável:</b> "."Eduardo Atene Silva");//$allTurmas->idFuncionarioResponsavel);
                    $card->setClassTextSub("d-none");
                    $card->setTextLeftButtonExist(false);;
                    $card->setImageExiste(true);
                    $card->setImagemBase($equipamento->image,$equipamento->imageNameFile);
                    $card->setColorCardbyClass("text-center");
                    $card->setColorButtonbyClass("d-none");
                    $div1->InsertRow('3',$card);
                }

                $div->InsertRow('12',$cardElements);
                $div->InsertRow('12',$div1);
                $div->show();
                ?>
                    <div class="modal fade" id="modalEquipamento" role="dialog" style="height: 100vh">
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
                Utility::InsertJavascriptFile("equipamento");

            break;
        }
        break;
    }

?>
