<?php
	$login  = Utility::getVariable("login");
	$password  = Utility::getVariable("password");
    $arrayAux = (array("login"=>$login,"pass"=>$password));
    $try = json_decode(logar($arrayAux));

    if($try->status == '200'){
        if($try->data->idUser==-1){
            $_SESSION['statusLogin'] = "fail";
            $error = "
                        <script>								
                                var url = location.protocol + '//' + location.host + location.pathname;
                                
                                action = function(){ window.location = url; }
                                setTimeout(action, 1);
                        </script>";
            echo $error;
        }
        else{
            Session::startSession($try->data->idUser);
            foreach ($try->data as $key => $value){
                $_SESSION[$key] = $value;
            }
            $ch='';

            //sucesso no login
            if(array_key_exists("idCliente",$_SESSION)){
                if($_SESSION["idCliente"]>0){ // Caso ele seja cliente
                    header("location: ./modalidades");
                }
                else{
                    $user = new Usuarioo($_SESSION['idUser']);
                    if($user->idTipoAcesso_id==1){ // Funcionario ADM
                        header("location: ./clientes");
                    }
                    if($user->idTipoAcesso_id==2){ // Cliente
                        header("location: ./treinos");
                    }
                    else if($user->idTipoAcesso_id==3){ // Funcionario Personal
                        header("location: ./modalidades");
                    }
                }
            }
            else{
                header("location: ./");
            }
        }
    }

    function logar($arrayAuxLogin){
	    $intReturn = 0;
        if($arrayAuxLogin == null) 
            $intReturn = -1;
        else{
            $try = Usuarioo::tryLogin($arrayAuxLogin);
            if($try==-1){
                $data = array("idUser"=>-1);
            }
            else{
                $user = new Usuarioo($try);
                $retorno = array();
                $querySql = "	SELECT u.id as idUser, i.nomeCompleto as nome,cp.id as idCliente, u.img AS file,cp.ativo,cp.dataInicio 
                                FROM ".Usuarioo::$database." u
                                JOIN ".InfoDadoss::$database." i ON i.id=u.idInfoDados
                                LEFT JOIN ".Cliente::$database." cp ON cp.idUsuario = u.id
                                WHERE u.id = ". $user->id."
                                ORDER BY cp.id DESC";
                $fetch = Connection::query($querySql);
                if(Connection::$numRows>0){
                    foreach($fetch as $rs){
                        $retorno = $rs;
                        break;
                    }
                }

                $retorno["idUser"]=($retorno["idUser"]);
                $retorno["id"] = $retorno["idUser"];
                $retorno["idCliente"]=($retorno["idCliente"]);
                $data = $retorno;
            }
        }
        switch ($intReturn) {
            case -1:
                $response['status']=400;
                $response['status_message']="Invalid Request";
                $response['data']= null;
                break;	
            case 0:
                $response['status']=200;
                $response['status_message']="Login Successful";
                $response['data']= $data;
            break;
        }
        return json_encode($response);
    }
