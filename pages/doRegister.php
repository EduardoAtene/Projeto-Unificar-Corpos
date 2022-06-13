<?php
require_once("./config/config.php");
    $nome  = Utility::getVariable("nome");
    $email = Utility::getVariable("email");
    $cpf  = Utility::getVariable("CPF");
    $telefone = Utility::getVariable("telefone");
    $dataNascimento  = Utility::getVariable("dataNascimento");
    $endereco  = Utility::getVariable("endereco");
    $login  = Utility::getVariable("login");
    $password  = Utility::getVariable("password");
    if(isset($nome)&&isset($login)&&isset($email)&&isset($password)){

        $response = Usuarioo::search(json_encode(array("login"=>$login)));
        $response = json_decode($response,true);
        if($response["status"]!='200'){
            //cria novo usuario
            $infoDados = array("email"=>$email,"nomeCompleto" => $nome,"CPF" => $cpf, "telefone" => $telefone,"dataNascimento"=>Utility::dateFormatToEU($dataNascimento),"endereco"=> $endereco);
            $idInfoDados = InfoDadoss::set(json_encode($infoDados));
            $dados = array("login"=>$login,"senha"=>sha1("kito".$password),"idTipoAcesso_id"=>2,"idInfoDados"=>$idInfoDados); //cria como perito
            $idUsuario = Usuarioo::set(json_encode($dados));
            Cliente::set(json_encode(array("idUsuario"=>$idUsuario, "dataInicio" => date("Y-m-d H:i:s"),"ativo" => 1)));

            $_SESSION['statusLogin'] = "UserCreated";
            $error = "
                        <script>								
                                var url = location.protocol + '//' + location.host + location.pathname;
                                
                                action = function(){ window.location = url; }
                                setTimeout(action, 1);
                        </script>";
            echo $error;
        }
        else{
            $user = new Usuarioo($response["data"][0]);
            $_SESSION['statusLogin'] = "alreadyHasUser";
            $error = "
                        <script>								
                                var url = location.protocol + '//' + location.host + location.pathname;
                                
                                action = function(){ window.location = url; }
                                setTimeout(action, 1);
                        </script>";
            echo $error;
            // Utility::showTimer(4500, "Seu CPF já possui um cadastro ativo. <br>Seu usuário é: <Br><b>".$user->login."</b><br>Se tiver esquecido a senha, favor usar o recurso de recuperar senha na página de login.");
        }
    }
