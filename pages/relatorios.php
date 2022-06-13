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
                case 6:
                    $all = true;
                    $response = Cliente::getAllClienteOrderName();
                    $data = array();
                    $template = array(  "CLIENTE" => "nomeCompleto",
                                        "CPF" => "CPF",
                                        "EMAIL"=>"email",
                                        "TELEFONE"=>"telefone",
                                        "DATA DE NASCIMENTO"=>"dataNascimento",
                                        "ENDEREÇO"=>"endereco");
                    $objPHPExcel = new PHPExcel();
                    $objPHPExcel->setActiveSheetIndex(0);
                    $dados = array("usuario"=>$_SESSION['nome']);
                    firularClientes($objPHPExcel,$dados);
                    $objPHPExcel->getActiveSheet()->fromArray(array_keys($template), NULL, 'A6');
                    $rowCount = 6;
                    $totalProcessos = count($response);
                    foreach($response as $nesimoProcesso=>$cliente){
                        $collumCount = 0;
                        $arraySomaPeriodoPago = array();
                        foreach($template as $key => $value){
                            $column = PHPExcel_Cell::stringFromColumnIndex($collumCount);
                            $row = $rowCount;
                            $cell = $column.$row;
                            //centralizar célula:
                            $objPHPExcel->getActiveSheet()->getStyle($cell)->applyFromArray(
                                array(
                                    'alignment' => array(
                                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                    )
                                )
                            ); 
                            switch($value){                                    
                                case "dataNascimento":
                                    if($cliente[$value] != "")
                                        $objPHPExcel->getActiveSheet()->setCellValueExplicit($cell,Utility::dateFormatToBR($cliente[$value]),PHPExcel_Cell_DataType::TYPE_STRING);
                                    $objPHPExcel->getActiveSheet()->getStyle($cell)->applyFromArray(
                                        array(
                                            'fill' => array(
                                                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                                'color' => array('rgb' => 'dbdbdb')
                                            ),
                                        )
                                    );
                                break;
                                case "nomeCompleto":                            
                                    $objPHPExcel->getActiveSheet()->setCellValueExplicit($cell,$cliente[$value],PHPExcel_Cell_DataType::TYPE_STRING);
                                    $objPHPExcel->getActiveSheet()->getStyle($cell)->applyFromArray(
                                        array(
                                            'borders' => array(
                                                'left' => array(
                                                    'style' => PHPExcel_Style_Border::BORDER_THICK,
                                                    'color' => array('rgb' => '045E7B')
                                                )
                                            )
                                        )
                                    ); 
                                break;
                                case "endereco":
                                    $objPHPExcel->getActiveSheet()->setCellValueExplicit($cell,$cliente[$value],PHPExcel_Cell_DataType::TYPE_STRING);
                                    $objPHPExcel->getActiveSheet()->getStyle($cell)->applyFromArray(
                                        array(
                                            'alignment' => array(
                                                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
                                            ),
                                            'fill' => array(
                                                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                                'color' => array('rgb' => 'dbdbdb')
                                            ),
                                            'borders' => array(
                                                'right' => array(
                                                    'style' => PHPExcel_Style_Border::BORDER_THICK,
                                                    'color' => array('rgb' => '000000')
                                                )
                                            )
                                        )
                                    ); 
                                break;
                                default:
                                    $objPHPExcel->getActiveSheet()->setCellValueExplicit($cell,$cliente[$value],PHPExcel_Cell_DataType::TYPE_STRING);
                                break;
                            }
                            $objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); 
                            $collumCount+=1;
                        }

                        $rowCount+=1;
                        if($nesimoProcesso==$totalProcessos-1){
                            $collumCount = 0;
                            for($i=0;$i<count($template);$i++){
                                $column = PHPExcel_Cell::stringFromColumnIndex($collumCount);
                                $row = $rowCount;
                                $cell = $column.$row;
                                $objPHPExcel->getActiveSheet()->getStyle($cell)->applyFromArray(
                                    array(
                                        'alignment' => array(
                                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                                        ),
                                        'borders' => array(
                                            'top' => array(
                                                'style' => PHPExcel_Style_Border::BORDER_THICK,
                                                'color' => array('rgb' => '000000')
                                            )
                                        )
                                    )
                                ); 
                                $collumCount+=1;
                            }
                        }
                    }

                    $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
                    ob_start();
                    $objWriter->save("php://output");
                    $xlsData = ob_get_contents();
                    ob_end_clean();
                    
                    $response =  array(
                            'status' => 200,
                            'fileName' => "Relatório do Cliente",
                            'file' => "data:application/vnd.ms-excel;base64,".base64_encode($xlsData)
                        );
                    
                    die(json_encode($response));

                break;
			}
		break;
        default:
        break;
    }
    Utility::InsertJavascriptFile("relatorios");
    
    function firularClientes($objPHPExcel,$dados){
        $objPHPExcel->getActiveSheet()->setShowGridlines(false);
        $objDrawing = new PHPExcel_Worksheet_Drawing();
        $objDrawing->setName('test_img');
        $objDrawing->setDescription('test_img');
        $objDrawing->setPath('../images/logo_extenso_sem_fundo.png');
        $objDrawing->setCoordinates('A1');                      
        //setOffsetX works properly
        $objDrawing->setOffsetX(50);
        $objDrawing->setOffsetY(2);
        //set width, height
        $objDrawing->setWidth(75);
        $objDrawing->setHeight(75);
        $objDrawing->setWorksheet($objPHPExcel->getActiveSheet());
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('C2','Extraído em: ',PHPExcel_Cell_DataType::TYPE_STRING);
        // $objPHPExcel->getActiveSheet()->setCellValueExplicit('C3','Cliente: ',PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('C3','Usuário: ',PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('D2',date('d/m/Y H:i:s'),PHPExcel_Cell_DataType::TYPE_STRING);
        // $objPHPExcel->getActiveSheet()->setCellValueExplicit('D3',$dados["cliente"],PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('D3',$dados["usuario"],PHPExcel_Cell_DataType::TYPE_STRING);

        //style: largura colunas
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(30); //dimensiona colunas
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20); //dimensiona colunas
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(40); //dimensiona colunas
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(25); //dimensiona colunas
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30); //dimensiona colunas
        $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(50); //dimensiona colunas
        $arrayColunas = array('A','B','C','D','E','F');

        $arrayRows = array(6,7);

        foreach ($arrayColunas as $col){
            foreach ($arrayRows as $row){
                $cell = $col.$row;#045E7B
                //var_dump($objPHPExcel->getActiveSheet()->getStyle($cell));
                $objPHPExcel->getActiveSheet()->getStyle($cell)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => '000000')
                        ),
                        'font'  => array(
                            'bold'  => true,
                            'color' => array('rgb' => 'FFFFFF'),
                            'size'  => 10,
                            'name'  => 'Calibri',
                        ),
                        'alignment' => array(
                            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                        ),
                        'borders' => array(
                            'allborders' => array(
                                'style' => PHPExcel_Style_Border::BORDER_THICK,
                                'color' => array('rgb' => '000000')
                            )
                        )
                    )
                ); 
            }
        }
    }
    function headerAjustest($objPHPExcel){
        // LAYOUT. Só para tampar burco pois isso me estressou '0'
        $objPHPExcel->getActiveSheet()->mergeCells('A5:A6');
        $objPHPExcel->getActiveSheet()->getStyle('A5:A6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER); 
        $objPHPExcel->getActiveSheet()->mergeCells('B5:B6');
        $objPHPExcel->getActiveSheet()->getStyle('B5:B6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('C5:C6');
        $objPHPExcel->getActiveSheet()->getStyle('C5:C6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('D5:D6');
        $objPHPExcel->getActiveSheet()->getStyle('D5:D6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('E5:E6');
        $objPHPExcel->getActiveSheet()->getStyle('E5:E6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('F5:G5');
        $objPHPExcel->getActiveSheet()->getStyle('F5:G5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('H5:H6');
        $objPHPExcel->getActiveSheet()->getStyle('H5:H6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('I5:I6');
        $objPHPExcel->getActiveSheet()->getStyle('I5:I6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('J5:J6');
        $objPHPExcel->getActiveSheet()->getStyle('J5:J6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('K5:K6');
        $objPHPExcel->getActiveSheet()->getStyle('K5:K6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('L5:L6');
        $objPHPExcel->getActiveSheet()->getStyle('L5:L6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('M5:M6');
        $objPHPExcel->getActiveSheet()->getStyle('M5:M6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('N5:N6');
        $objPHPExcel->getActiveSheet()->getStyle('N5:N6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        // $objPHPExcel->getActiveSheet()->mergeCells('J5:N5');
        // $objPHPExcel->getActiveSheet()->getStyle('J5:N5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('O5:O6');
        $objPHPExcel->getActiveSheet()->getStyle('O5:O6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('P5:P6');
        $objPHPExcel->getActiveSheet()->getStyle('P5:P6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('Q5:Q6');
        $objPHPExcel->getActiveSheet()->getStyle('Q5:Q6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('R5:R6');
        $objPHPExcel->getActiveSheet()->getStyle('R5:R6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->mergeCells('S5:T5');
        $objPHPExcel->getActiveSheet()->getStyle('S5:T5')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        // $objPHPExcel->getActiveSheet()->mergeCells('O5:O6');
        // $objPHPExcel->getActiveSheet()->getStyle('O5:O6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        // $objPHPExcel->getActiveSheet()->mergeCells('P5:P6');
        // $objPHPExcel->getActiveSheet()->getStyle('P5:P6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        // $objPHPExcel->getActiveSheet()->mergeCells('Q5:Q6');
        // $objPHPExcel->getActiveSheet()->getStyle('Q5:Q6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        // $objPHPExcel->getActiveSheet()->mergeCells('R5:R6');
        // $objPHPExcel->getActiveSheet()->getStyle('R5:R6')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('A5',"PRIORIDADE",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('B5',"Nº PROCESSO",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('C5',"DATA EXTRAÇÃO",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('D5',"AUTOR",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('E5',"RÉU",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('F5',"ORGANIZADOR DE RECEBIMENTO",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('H5',"PGTO PELO TRE (Res 66)",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('I5',"PGTO AO PERITO",PHPExcel_Cell_DataType::TYPE_STRING);
        // $objPHPExcel->getActiveSheet()->setCellValueExplicit('J5',"LEITURA SOMENTE NOS STATUS DE PAGAMENTO MÉDIO E FRIO",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('J5',"OUTROS PAGAMENTOS",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('K5',"ACORDO",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('L5',"BUSCA PATRIMONIAL",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('M5',"ARQUIVO",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('N5',"ÚLTIMO DESPACHO",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('O5',"SUGESTÃO PARA CONSULTORIA",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('P5',"OBJETIVOS DA CONSULTORIA",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('Q5',"REGISTROS PREVISTOS",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->getStyle('J5')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('K5')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('L5')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('M5')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('N5')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('O5')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('P5')->getAlignment()->setWrapText(true);
        $objPHPExcel->getActiveSheet()->getStyle('Q5')->getAlignment()->setWrapText(true);

        $objPHPExcel->getActiveSheet()->setCellValueExplicit('R5',"CONSULTORIA",PHPExcel_Cell_DataType::TYPE_STRING);
        $objPHPExcel->getActiveSheet()->setCellValueExplicit('S5',"CONSULTORIA",PHPExcel_Cell_DataType::TYPE_STRING);
        // $objPHPExcel->getActiveSheet()->setCellValueExplicit('O5',"DATA ENVIO CONSULTORIA",PHPExcel_Cell_DataType::TYPE_STRING);
        // $objPHPExcel->getActiveSheet()->setCellValueExplicit('P5',"DATA RESOLUÇÃO CONSULTORIA",PHPExcel_Cell_DataType::TYPE_STRING);
        // $objPHPExcel->getActiveSheet()->setCellValueExplicit('Q5',"PETIÇÃO",PHPExcel_Cell_DataType::TYPE_STRING);
        // $objPHPExcel->getActiveSheet()->setCellValueExplicit('R5',"COMENTÁRIOS CONSULTORIA",PHPExcel_Cell_DataType::TYPE_STRING);

        
        //$objPHPExcel->getActiveSheet()->setCellValueExplicit('A5',"Prioridade",PHPExcel_Cell_DataType::TYPE_STRING);

    }
?>